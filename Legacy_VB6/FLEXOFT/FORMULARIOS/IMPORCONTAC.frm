VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form IMPORCONTAC 
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
      _Version        =   393216
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
      Height          =   4245
      Left            =   105
      TabIndex        =   16
      ToolTipText     =   "Importación Definitiva"
      Top             =   3990
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   41
         Top             =   420
         Width           =   10215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -20
            ScaleHeight     =   1965
            ScaleWidth      =   10215
            TabIndex        =   43
            Top             =   0
            Width           =   10215
            Begin VB.PictureBox Picture3 
               Height          =   1950
               Left            =   -20
               ScaleHeight     =   1890
               ScaleWidth      =   0
               TabIndex        =   48
               Top             =   0
               Width           =   60
            End
            Begin VB.PictureBox Picture7 
               Height          =   270
               Left            =   0
               ScaleHeight     =   210
               ScaleWidth      =   1.1994e5
               TabIndex        =   45
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
                  TabIndex        =   47
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
                  TabIndex        =   46
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
               ItemData        =   "IMPORCONTAC.frx":0000
               Left            =   0
               List            =   "IMPORCONTAC.frx":0002
               TabIndex        =   44
               Top             =   250
               Width           =   1.3e5
            End
         End
         Begin VB.PictureBox Picture10 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   42
            Top             =   0
            Width           =   60
         End
      End
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
         ItemData        =   "IMPORCONTAC.frx":0004
         Left            =   120
         List            =   "IMPORCONTAC.frx":0006
         TabIndex        =   26
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
         Picture         =   "IMPORCONTAC.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   25
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
         TabIndex        =   24
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
         Picture         =   "IMPORCONTAC.frx":0312
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
         Bindings        =   "IMPORCONTAC.frx":061C
         Height          =   3750
         Left            =   150
         OleObjectBlob   =   "IMPORCONTAC.frx":062C
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
         TabIndex        =   31
         Top             =   350
         Width           =   4110
         Begin VB.VScrollBar VScroll1 
            Height          =   2050
            Left            =   3830
            TabIndex        =   36
            Top             =   0
            Width           =   225
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFF0&
            BorderStyle     =   0  'None
            Height          =   4560
            Left            =   0
            ScaleHeight     =   4560
            ScaleWidth      =   4740
            TabIndex        =   32
            Top             =   0
            Width           =   4740
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Listas de Precios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   240
               TabIndex        =   49
               Top             =   3700
               Width           =   3420
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Movimientos de Stock Valorizados"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   240
               TabIndex        =   40
               Top             =   4200
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Cartera de Valores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   240
               TabIndex        =   39
               Top             =   3450
               Width           =   3420
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Composición de Deuda Proveedores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   240
               TabIndex        =   38
               Top             =   3200
               Width           =   3420
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Composición de Deuda Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   240
               TabIndex        =   37
               Top             =   2950
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Numeros Originales Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   255
               TabIndex        =   35
               Top             =   2500
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Plan de Cuentas Contables"
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
               Index           =   7
               Left            =   255
               TabIndex        =   34
               Top             =   2025
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Planilla Relevamiento de Revistas"
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
               TabIndex        =   33
               Top             =   105
               Width           =   3735
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   4000
               Y2              =   4000
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
         TabIndex        =   28
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
            TabIndex        =   29
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
      Picture         =   "IMPORCONTAC.frx":0FDE
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
      Picture         =   "IMPORCONTAC.frx":12E8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   30
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORCONTAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CANCELIMP%
'
Private Sub Combo1_Change()
   '
   Data1.RecordSource = Combo1.Text
   Data1.Refresh
   Data1.Recordset.MoveFirst
   '
End Sub

Private Sub Combo1_Click()
   '
   Data1.RecordSource = Combo1.Text
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
   Combo1.Text = ""
   On Error Resume Next
   Data1.Refresh
   '
   CATAB% = Data1.Database.TableDefs.Count
   For TTB% = 1 To CATAB%
     Combo1.AddItem Data1.Database.TableDefs(TTB% - 1).Name
     If TTB% = 1 Then
       Combo1.Text = Data1.Database.TableDefs(TTB% - 1).Name
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
     List2.AddItem " 11  Kilos por Unidad           Num      0 - 10^8"
     List2.AddItem " 12  Kilos por Metro            Num      0 - 10^8"
     List2.AddItem " 13  Costo Unitario             Num      0 - 10^8"
     List2.AddItem " 14  Moneda de Costeo           Num      1 - 255"
     List2.AddItem " 15  Fecha de Costeo            Date"
     List2.AddItem " 16  Numero Proveedor Habitual  Num      0 - 32767"
     List2.AddItem " 17  Grupo Contable             Num      0 - 32767"
     List2.AddItem " 18  Stock Actual               Num      +/- 999999"
     ENCABE$ = "Código           C-Int  Cod.Alternativo          Descripcion                                                      Descripción Complementaria                                       Atributo A                       Atributo B                         TM UM         St.Min      Lote Ec Kilos/Un  Kilos/Mt    Costo Unit   MC Fecha        Pr-H     GC       Stock"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
   Case 1
     List2.AddItem "  1  Codigo de Conjunto         String   <= 15 car"
     List2.AddItem "  2  Codigo del Componente      String   <= 15 car"
     List2.AddItem "  3  Uso / Cantidad Componente  Num      10^-5 - 10^6"
     List2.AddItem "  4  Unidades                   Num      10^-5 - 10^6"
     List2.AddItem "  5  Largo de corte (mm)        Num      10^-5 - 10^6"
     List2.AddItem "  6  Plazo Disposicion          Num      1 - 255"
     List2.AddItem "  7  Referencia                 String"
     Label8 = "Código Conj.              Código Elem.              Descripcion                         Uso/Cant     Unidades        Corte   PD   Referencia"
     Label82 = Space(300)
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
     List2.AddItem " 17  Descuento Standard         String   <= 24 car"
     List2.AddItem " 18  L.Entrega - Calle + Nro    String   <= 48 car"
     List2.AddItem " 19  L.Entrega - Localidad      String   <= 48 car"
     List2.AddItem " 20  Transportista              String   <= 48 car"
     List2.AddItem " 21  Grupo Divisional           Num      0 - 9999"
     List2.AddItem " 22  Percepcion I/Brutos        Num      0 / 1"
     '
     ENCABE$ = " Cta. Razón Social                                    Domicilio                                       " & _
     "C.Postal    Localidad                       Teléfono                Fax                     Contacto                CUIT             Gp.  Cod.Cta       V.  Pr  PI  CP     Lim.Cre  " & _
     "Descuento                 Lugar de Entrega                                Localidad                                       Transportista                                    Gp.Div.  " & _
     "IB"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
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
     List2.AddItem " 16  Retencion I/Brutos         Num      0 - 255"
     ENCABE$ = "Cta.  Razón Social                                    Domicilio                                       C.Postal    Localidad                       " + _
     "Teléfono                Fax                     Contacto                CUIT            T.   Cod.Cta.    Pr. IVA  Cta. Madre    N.Fantasia          IB"
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
   Case 13
     '\\\OT-5-0641-FG-26/09/05///
     List2.AddItem "  1  Número de Lista            Nume     1 - 255"
     List2.AddItem "  2  Código de Artículo         String   <= 15 car"
     List2.AddItem "  3  Número Original            String   <= 16 car"
     List2.AddItem "  4  Precio Unitario            Num      0 - 10^8"
     Label8 = "Nro. Código            Descripción                                                       Nro.Original        P.Unitario"
     '\\\OT-5-0641-FG-FIN///
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
   
   Dim RAEMP$()
   CAEMP& = ULTREG&("BSEGCOM")
   ReDim Preserve RAEMP$(CAEMP& + 1)
   For U& = 1 To ULTREG&("BSEGCOM")
     X$ = REGLEIDO$("BSEGCOM", U&)
     RAEMP$(U&) = UCase$(TRIM$(Mid$(X$, 3, 62)))
   Next U&
   '
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(313)
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
     CODEX$ = TRIM$(DBGrid1.Text)
     'VERIFICAR SI EXISTENTE
     For KU& = 1 To CAEMP&
       If UCase$(RAEMP$(KU&)) = CODEX$ Then
         List1.AddItem CODEX$
         GoTo 19
       End If
     Next KU&
     '
12   Call REPLA(TTXX$, CODEX$, 1, 64)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, DEDEX$, 65, 48)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.Text, 8)
     Call REPLA(TTXX$, DOMI$, 113, 8)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, CODPOS$, 121, 32)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, LOCA$, 153, 24)
     '
     DBGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, TELEF$, 177, 32)
     '
     DBGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, FAX$, 209, 32)
     '
     DBGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, CONTACTO$, 241, 48)
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, CUIT$, 289, 8)
     '
     DBGrid1.Col = 9                     ' GRUPO DE CLIENTE
     GRUPOCLI$ = TRIM$(DBGrid1.Text)
     Call REPLA(TTXX$, GRUPOCLI$, 297, 16)
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
   DBGrid1.RowHeight = HROW   ' RESTAURA ALTURA ANTERIOR
   '
End Sub
'
Sub IMPORMAS()
   '
   IMAX% = 0
   For U& = 1 To ULTREG&("BSEGCOM")
     X$ = REGLEIDO$("BSEGCOM", U&)
     If CVI(Left$(X$, 2)) > IMAX% Then
       IMAX% = CVI(Left$(X$, 2))
     End If
   Next U&
   '
   FIN& = List4.ListCount
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     A$ = List4.List(U& - 1)
     EMPRA$ = TRIM$(Left$(A$, 64))
     CANU$ = TRIM$(Mid$(A$, 65, 48))
     CPOS$ = TRIM$(Mid$(A$, 113, 8))
     LOCA$ = TRIM$(Mid$(A$, 121, 32))
     TELE$ = TRIM$(Mid$(A$, 153, 24))
     DIREL$ = TRIM$(Mid$(A$, 177, 32))
     SITE$ = TRIM$(Mid$(A$, 209, 32))
     PRIP$ = TRIM$(Mid$(A$, 241, 48))
     FABIN$ = TRIM$(Mid$(A$, 289, 8))
     ORIGE$ = TRIM$(Mid$(A$, 297, 16))
     GoSub ALTANUE
   Next U&
   Call CIERRARCH("BSEGCOM")
   '
Exit Sub
'
ALTANUE:
   IMAX% = IMAX% + 1
5  IFICHA% = IMAX%
   RELE$ = REGBLAN$("BSEGCOM")
   '
   Call REPLA(RELE$, MKI$(IFICHA%), 1, 2)
   Call REPLA(RELE$, EMPRA$, 3, 62)      ' empresa
   Call REPLA(RELE$, DIREL$, 65, 64)     ' DIRECCION ELECTRONICA
   Call REPLA(RELE$, CANU$, 129, 48)   ' CALLE Y NUMERO
   Call REPLA(RELE$, CPOS$, 177, 8)    ' C.POSTAL
   Call REPLA(RELE$, LOCA$, 185, 32)   ' LOCALIDAD
   Call REPLA(RELE$, "ARGENTINA", 249, 24)   ' PAIS
   '
   Call REPLA(RELE$, ORIGE$, 273, 24)   ' ORIGEN
   Call REPLA(RELE$, FABIN$, 297, 24)    ' RAMO
   Call REPLA(RELE$, PRIP$, 321, 48)    ' PRINC.PRODUCTO
   Call REPLA(RELE$, "", 369, 8) ' provincia particular
   Call REPLA(RELE$, TELE$, 377, 24)   ' TELEFONO
   Call REPLA(RELE$, "", 401, 24) ' telef.partic.
   Call REPLA(RELE$, "", 425, 24) ' telef2.partic
   '
   Call REPLA(RELE$, SITE$, 737, 64) ' WEB-SITE
   
   Call REGAPP("BSEGCOM", RELE$)
   '
89 Return
'


End Sub

