VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ARCHSIMT 
   BackColor       =   &H00D0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S.Imp.Temp. - Archivos Maestros"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   435
      Left            =   6720
      TabIndex        =   10
      Top             =   945
      Width           =   1380
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\SABO_WIN\TIPIFICA\ANEXO ACELERANTES.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   540
      Left            =   5145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "'ANEXO ACELERANTES'$"
      Top             =   1890
      Width           =   1905
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7770
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   435
      Left            =   8610
      TabIndex        =   8
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command6 
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
      Height          =   690
      Left            =   5145
      Picture         =   "ARCHSIMT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5145
      Picture         =   "ARCHSIMT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1050
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0D0&
      Caption         =   "Planillas de Tipificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4845
      Begin VB.CommandButton Command4 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   5
         Top             =   1785
         Width           =   1140
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Configurar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2415
         TabIndex        =   4
         Top             =   1785
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   3
         Top             =   1785
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   1785
         Width           =   1140
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
         Height          =   1230
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   4600
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ARCHSIMT.frx":0454
      Height          =   49995
      Left            =   210
      OleObjectBlob   =   "ARCHSIMT.frx":0464
      TabIndex        =   9
      Top             =   2625
      Width           =   11000
   End
End
Attribute VB_Name = "ARCHSIMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command7_Click()
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Conectar Archivo Planilla de Producción"
       'deshabilitado porque si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Planillas Excel (*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error Resume Next
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   Data1.DatabaseName = CommonDialog1.filename
   '
End Sub

Private Sub Command8_Click()
   FORTIPIF.Show 1
End Sub

