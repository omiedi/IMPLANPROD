VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IMPPED07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Pedidos Desde Planillas Excel"
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
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   47
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0E0&
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
      TabIndex        =   15
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
         ItemData        =   "IMPPED07.frx":0000
         Left            =   120
         List            =   "IMPPED07.frx":0002
         TabIndex        =   28
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
         Picture         =   "IMPPED07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         Enabled         =   0   'False
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
         Picture         =   "IMPPED07.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1800
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1835
         Left            =   10360
         TabIndex        =   19
         Top             =   700
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   18
         Top             =   2310
         Width           =   10250
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   16
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
            ItemData        =   "IMPPED07.frx":0618
            Left            =   0
            List            =   "IMPPED07.frx":061A
            TabIndex        =   17
            Top             =   240
            Width           =   40000
         End
         Begin VB.PictureBox Picture7 
            Height          =   270
            Left            =   0
            ScaleHeight     =   210
            ScaleWidth      =   89940
            TabIndex        =   33
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
               TabIndex        =   35
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
               TabIndex        =   34
               Top             =   0
               Width           =   120
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   20
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
         TabIndex        =   29
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0E0&
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
      Begin MSFlexGridLib.MSFlexGrid DBGrid1 
         Height          =   3750
         Left            =   150
         TabIndex        =   49
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6615
         _Version        =   393216
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
      BackColor       =   &H00C0E0E0&
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
      BackColor       =   &H00C0E0E0&
      Caption         =   "Datos del Cliente"
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
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2385
         MaxLength       =   16
         TabIndex        =   44
         Top             =   2160
         Width           =   1920
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   3315
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   960
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   240
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   4040
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
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
         Left            =   225
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1065
         Width           =   4065
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   225
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1425
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1065
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1425
         Width           =   3220
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00E0E0E0&
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
         Left            =   240
         MaxLength       =   40
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1785
         Width           =   4050
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Compra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         TabIndex        =   45
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2520
         TabIndex        =   43
         Top             =   405
         Width           =   750
      End
      Begin VB.Label Label26 
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
         Height          =   300
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
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
      Begin VB.CommandButton Command7 
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
         Height          =   550
         Left            =   10920
         Picture         =   "IMPPED07.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Inicia Pre-Importación"
         Top             =   360
         Width           =   550
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   30
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
            TabIndex        =   31
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
         Width           =   2295
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
      Enabled         =   0   'False
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
      Picture         =   "IMPPED07.frx":0926
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
      Picture         =   "IMPPED07.frx":0C30
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   32
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "IMPPED07.frx":0D7A
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "IMPPED07.frx":0FAE
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label11 
      Caption         =   "NRO. LIPRE"
      Height          =   375
      Left            =   2640
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPPED07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private HOJA As Excel.Worksheet
Private Sub Combo1_Change()
   '
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.TEXT = Libro.Sheets(KKI%).Name Then
        Set HOJA = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set HOJA = Libro.Sheets(1)
   DBGrid1.Clear
   DBGrid1.Rows = 1

   '
End Sub

Private Sub Combo1_Click()
   '
   'Set objexcel = New Excel.Application
   'Set Libro = objexcel.Workbooks.Open(Label1.Caption)
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.TEXT = Libro.Sheets(KKI%).Name Then
        Set HOJA = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set HOJA = Libro.Sheets(1)
   DBGrid1.Clear
   DBGrid1.Rows = 1   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")

   On Error Resume Next
   Set HOJA = Nothing
   Libro.Close: Set Libro = Nothing
   objexcel.Close: Set objexcel = Nothing
   On Error GoTo 0
   Unload Me
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   Call IMPORPED
   Screen.MousePointer = 1
   Frame5.Visible = False
   DoEvents
   
   Call CIERRARCH("*.*")

   On Error Resume Next
   Set HOJA = Nothing
   Libro.Close: Set Libro = Nothing
   objexcel.Close: Set objexcel = Nothing
   On Error GoTo 0
   Unload Me
   
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.*)"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   Openforms = DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.FileName)
'   On Error Resume Next
'   If Err Then
'      'GoTo 98
'      ERNU$ = TRIM$(Str$(Err.Number))
'      On Error GoTo 0
'      If ERNU$ = "3125" Then GoTo 98    ' falta revisar - 25/01/2008
'      MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") al Abrir la Hoja de Cálculo Especificada. Imposible Completar la Petición. Consulte."
'      MsgBox MENSA$
'      GoTo 99
'   End If
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
   Frame4.Visible = False
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
   URECORD& = DBGrid1.Rows   '   Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DBGrid1.Height
   'DbGrid1.Height = (URECORD& + 10) * DbGrid1.RowHeight
   'Data1.Recordset.MoveFirst
   '
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   Call PREIMPOPED
  '
  'Libro.Close
  'objexcel.Quit
  '
  If MOSTRARMENSAJE = True Then
     Call COMUNI("Se Ha Completado la Pre-Importación.")
  End If
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
   Command2.Enabled = True
End Sub


Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub


Private Sub Command7_Click()

   Screen.MousePointer = 11
   CANFIL& = 10000   ' Hoja.Rows
   CANCOL& = List2.ListCount    ' Hoja.Columns
   JK& = 0
   CONTAERR& = 0
   ultimafila& = 0
   
   
   For i& = 1 To CANFIL&
     YAGRELI% = 0
     For J& = 1 To CANCOL&
       valor = HOJA.Cells(i&, J&)
       If TRIM$(CStr(valor)) <> "" Then
         If YAGRELI% < 1 Then
           YAGRELI% = 1
           JK& = JK& + 1
         End If
         If JK& > DBGrid1.Rows Then DBGrid1.Rows = DBGrid1.Rows + 1
         If JK& < 1 Then Call COMUNI("Imposible Continuar.\Falta Encabezado en Alguna de las Columnas"): Screen.MousePointer = 1: Exit Sub
         
         DBGrid1.Row = JK& - 1
       Else
         If YAGRELI% = 0 Then
            If ultimafila& < 1 Then
                ultimafila& = DBGrid1.Row + 1
                If ultimafila& >= DBGrid1.Rows Then DBGrid1.Rows = DBGrid1.Rows + 1
            End If
            DBGrid1.Row = ultimafila&
            YAGRELI% = 1
         End If
       End If
         If J& > DBGrid1.Cols Then DBGrid1.Cols = DBGrid1.Cols + 1
       DBGrid1.COL = J& - 1
       '
       If DBGrid1.COL = 4 And JK& > 1 Then
            If IsDate(valor) = False Then
                valor = FECHATEX$(HOY(HO$))
            End If
       End If
       On Error Resume Next
       DBGrid1.TEXT = valor
       'CAPTURA DE ERROR DE DATOS NO VALIDOS DE PLANILLA EXCEL
       If Err > 0 Then
           CONTAERR& = CONTAERR& + 1
          'GUARDO EL DATO DE FILA Y COLUMNA DE LA PRIMER OCURRENCIA DE ERROR
          If CONTAERR& = 1 Then Fil& = i&: COL& = J&
          DBGrid1.TEXT = ""
       End If
       On Error GoTo 0
       
       '
       Label9.Font = DBGrid1.Font
       Label9.FontSize = DBGrid1.CellFontSize
       Label9.FontBold = DBGrid1.CellFontBold
       Label9.Caption = CStr(valor)
       If Label9.Width > DBGrid1.ColWidth(J& - 1) Then
         DBGrid1.ColWidth(J& - 1) = Label9.Width
       End If
     Next J&
     If i& > JK& + 256 Then
        GoTo 99
     End If
   Next i&
   '
99 Screen.MousePointer = 1
   '
   If ultimafila& > 0 Then DBGrid1.RemoveItem (ultimafila&)
   '
   'MENSAJE SI SE PRODUJO ERROR CAPTURADO DE DATO NO VALIDO
   If CONTAERR& > 0 Then
     Call MENSERR(24, "Se Han Detectado Uno o Más Datos No Válidos\ Primer Ocurrencia Fila: " & Fil& & " Columna: " & COL& & "\El o Los Mismos Fueron Cambiados por Texto Vacío \Puede Continuar Normalmente o Modificar La Planilla de Importación\ Y Realizar Nuevamente la Importación")
   End If
   '
   Call COMUNI("Origen de Datos Establecido.")
   Command4.Enabled = True
   '
End Sub


Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   Call CARGADET
    Call APLICACIONSKINS(Me)
   '
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

 Sub CARGADET()
   List1.Clear
   List2.Clear
   List4.Clear
   Frame5.Visible = False
   Label1 = ""
   Combo1.Clear
'   NARIM = TRIM$(Str$(Index))
     List2.AddItem "  1  Codigo de Articulo         String   <= 16 car"
     List2.AddItem "  2  Codigo de Cliente          String   <= 28 car"
     List2.AddItem "  3  Descripcion                String   <= 64 car"
     List2.AddItem "  4  Cantidad                   Num      0 - 10^6"
     List2.AddItem "  5  Fecha de Entrega           Date"
     List2.AddItem "  6  Precio Unitario            Num      0 - 10^8"
     Label8 = "Codigo           Cliente                       Descripcion                                                           Cantidad  F.Entrega         Precio"
     '
End Sub

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub Label26_gotfocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub label1_Change()
  DBGrid1.Clear
  DBGrid1.Rows = 1
  Combo1.Clear
  Combo1.TEXT = ""
  '
  If Label1.Caption <> "" Then
     'Dim objexcel As Excel.Application
     Set objexcel = New Excel.Application
     Set Libro = objexcel.Workbooks.Open(Label1.Caption)
     '
     CANOJ% = Libro.Sheets.Count
     For KKI% = 1 To CANOJ%
        Combo1.AddItem Libro.Sheets(KKI%).Name
        If KKI% = 1 Then
           Combo1.TEXT = Libro.Sheets(KKI%).Name
        End If
     Next KKI%
     Call Combo1_Click
     On Error Resume Next
     'Libro.Close
     'Set Libro = Nothing
     On Error GoTo 0
  End If

End Sub

Private Sub Text10_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text12_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text2_gotfocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text3_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text4_GotFocus()
  '
  Command22.SetFocus
  '
End Sub

Private Sub Text5_GotFocus()
  '
  Command22.SetFocus
  '
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

'Sub PREIMPORMAS()
'   '
'   If DBGrid1.Columns.Count < 10 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   Screen.MousePointer = 11
'   URECORD& = Data1.Recordset.RecordCount
'   ATX% = 0
'   '
'   With Data1
'     .Recordset.MoveFirst
'     U& = 0
'     Do While Not .Recordset.EOF
'        U& = U& + 1
'        Call TRACE(20, U&, URECORD&)
'        TTXX$ = Space$(254)
'        CODEX$ = ""
'        On Error Resume Next
'        CODEX$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
'        On Error GoTo 0
'        If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
'        '
'        CDXI% = CODINT%(CODEX$)
'        If CDXI% > 0 Then
'          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
'          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
'          ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
'          If ATX% <= 2 Then GoTo 19 ' SALTEAR
'        End If
'        '
'12      Call REPLA(TTXX$, CODEX$, 1, 16)
'        VALO$ = AJUSTD$(Str$(CDXI%), 6)
'        Call REPLA(TTXX$, VALO$, 17, 6)
'        '
'        DEDEX$ = ""
'        On Error Resume Next
'        DEDEX$ = TRIM$(.Recordset.Fields(1))
'        On Error GoTo 0
'        If Len(DEDEX$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
'        Call REPLA(TTXX$, DEDEX$, 25, 64)
'        '
'        VALO$ = ""
'        On Error Resume Next
'        VALO$ = AJUSTD$(.Recordset.Fields(2), 4)
'        On Error GoTo 0
'        If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Tipo de Material no Procesable"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
'        Call REPLA(TTXX$, VALO$, 91, 4)
'        '
'        UME$ = "U."
'        On Error Resume Next
'        UME$ = TRIM$(.Recordset.Fields(3))
'        On Error GoTo 0
'        Call REPLA(TTXX$, UME$, 97, 4)
'        '
'        valor = 0
'        On Error Resume Next
'        valor = CDbl(.Recordset.Fields(4))
'        On Error GoTo 0
'        VALO$ = FORMATNUM$(valor, "F12.1")
'        If XVALO(VALO$) < 0 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Minimo no Procesable"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
'        Call REPLA(TTXX$, VALO$, 101, 12)
'        '
'        valor = 0
'        On Error Resume Next
'        valor = CDbl(.Recordset.Fields(5))
'        On Error GoTo 0
'        VALO$ = FORMATNUM$(valor, "F12.1")
'        If XVALO(VALO$) < 0 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Lote de Reposicion no Procesable"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
'        Call REPLA(TTXX$, VALO$, 113, 12)
'        '
'        valor = 0
'        On Error Resume Next
'        valor = CDbl(.Recordset.Fields(6))
'        On Error GoTo 0
'        VALO$ = FORMATNUM$(valor, "F12.4")
'        Call REPLA(TTXX$, VALO$, 125, 12)
'        '
'        On Error Resume Next
'        VALO$ = AJUSTD$(.Recordset.Fields(7), 4)
'        On Error GoTo 0
'         If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Moneda de Costeo no Procesable"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
'        Call REPLA(TTXX$, VALO$, 137, 4)
'        '
'        VALO$ = ""
'        On Error Resume Next
'        VALO$ = Format$(.Recordset.Fields(8), "dd/mm/yy")
'        On Error GoTo 0
'        Call REPLA(TTXX$, VALO$, 143, 8)
'        '
'        valor = 0
'        On Error Resume Next
'        valor = CDbl(.Recordset.Fields(9))
'        On Error GoTo 0
'        If valor < 0 Or valor > 32766 Then valor = 0
'        VALO$ = FORMATNUM$(valor, "F8.0")
'        Call REPLA(TTXX$, VALO$, 151, 8)
'        '
'        valor = 0
'        On Error Resume Next
'        valor = CDbl(.Recordset.Fields(10))
'        On Error GoTo 0
'        If valor < (-9999999) Or valor > 9999999 Then valor = 0
'        VALO$ = FORMATNUM$(valor, "F10.1")
'        Call REPLA(TTXX$, VALO$, 161, 10)
'        '
'        List4.AddItem TTXX$
'        Label6 = TRIM$(Str$(List4.ListCount))
'        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'          TEPRUE = RTrim(TTXX$ + "*")
'        End If
'        On Error Resume Next
'        List4.TopIndex = List4.ListCount - 7
'        On Error GoTo 0
'        DoEvents
'19   .Recordset.MoveNext
'     Loop
'   End With
'   '
'20 Screen.MousePointer = 1
'   '
'End Sub
'
'Sub PREIMPORSTRU()
'   '
'   If DBGrid1.Columns.Count < 4 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   Screen.MousePointer = 11
'   URECORD& = Data1.Recordset.RecordCount
'   ATX% = 0: CODEXA$ = ""
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(254)
'     DBGrid1.Row = U& - 1
'     DBGrid1.COL = 0
'     CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
'     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
'     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
'     '
'     CODEXA$ = CODEX$
'     Call LEEEXPLO(CODEX$, 1)
'     If CAIT% > 0 And ATX% = 2 Then GoTo 19
'     '
'     'ATX% = 3
'     If CAIT% > 0 Then
'         ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
'       Else
'         GoTo 12
'     End If
'     '
'10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
'     '
'12   Call REPLA(TTXX$, CODEX$, 1, 16)
'     '
'     DBGrid1.COL = 1
'     COELE$ = TRIM$(DBGrid1.TEXT)
'     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento Inexistente o No Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, COELE$, 19, 16)
'     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 37, 32)
'     '
'     DBGrid1.COL = 2                     ' USO UNITARIO
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     VALO$ = FORMATNUM$(valor, "F12.4")
'     Call REPLA(TTXX$, VALO$, 69, 12)
'     '
'     DBGrid1.COL = 3                     ' PLAZO DISPOSICION
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     VALO$ = FORMATNUM$(valor, "I4")
'     Call REPLA(TTXX$, VALO$, 81, 4)
'     '
'     List4.AddItem TTXX$
'     Label6 = TRIM$(Str$(List4.ListCount))
'     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'       TEPRUE = RTrim(TTXX$ + "*")
'     End If
'     On Error Resume Next
'     List4.TopIndex = List4.ListCount - 7
'     On Error GoTo 0
'     DoEvents
'     '
'19 Next U&
'   Screen.MousePointer = 1
'   '
'20 End Sub
'
Sub PREIMPORUTS()
   '
'   If DBGrid1.Columns.Count < 12 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   Screen.MousePointer = 11
'   URECORD& = Data1.Recordset.RecordCount
'   ATX% = 0: CODEXA$ = "": CODEYA$ = ""
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(254)
'     DBGrid1.Row = U& - 1
'     DBGrid1.COL = 0
'     CODEX$ = TRIM$(DBGrid1.TEXT)
'     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
'     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
'     '
'     CODEXA$ = CODEX$
'     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(CODEX$)))
'     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
'     '
'     'ATX% = 3   ' SOBREESCRIBIR
'     If Len(RFF$) > 0 Then
'         ATX% = COMALTER%("Secuencia Operaciones de Código '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
'       Else
'         GoTo 12
'     End If
'     '
'10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
'     '
'12   Call REPLA(TTXX$, CODEX$, 1, 16)
'     If CODEX$ <> CODEYA$ Then NOPE% = 0
'     CODEYA$ = CODEX$
'     '
'     DBGrid1.COL = 1
'     DEOPE$ = TRIM$(DBGrid1.TEXT)
'     If DEOPE$ = "" Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion de la Operacion"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     NOPE% = NOPE% + 10
'     VALO$ = FORMATNUM$(NOPE%, "I4")
'     Call REPLA(TTXX$, VALO$, 17, 4)
'     Call REPLA(TTXX$, DEOPE$, 23, 48)
'     '
'     DBGrid1.COL = 2                     ' MAQUINA
'     CMAQUI$ = TRIM$(DBGrid1.TEXT)
'     If CMAQUI$ <> "" Then
'       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina '" + CMAQUI$ + "' Inexistente o No Valido"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'     End If
'     Call REPLA(TTXX$, CMAQUI$, 73, 6)
'     '
'     DBGrid1.COL = 3                     ' HORAS PREPARACION
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     VALO$ = FORMATNUM$(valor, "F10.3")
'     Call REPLA(TTXX$, VALO$, 79, 10)
'     '
'     DBGrid1.COL = 4                     ' HORAS STANDARD FIJO
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     VALO$ = FORMATNUM$(valor, "F10.3")
'     Call REPLA(TTXX$, VALO$, 89, 10)
'     '
'     DBGrid1.COL = 5                     ' MINUTOS STANDARD VARIABLE
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     VALO$ = FORMATNUM$(valor, "F12.5")
'     Call REPLA(TTXX$, VALO$, 99, 12)
'     '
'     DBGrid1.COL = 6                     ' LOTE STANDARD UNIDADES
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     If valor < 1 Then valor = 1
'     VALO$ = FORMATNUM$(valor, "F10.1")
'     Call REPLA(TTXX$, VALO$, 111, 10)
'     '
'     DBGrid1.COL = 7                     ' % DEM.DISTRIBUIDA
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     If valor < 0 Then valor = 0
'     On Error GoTo 0
'     VALO$ = FORMATNUM$(valor, "F8.1")
'     Call REPLA(TTXX$, VALO$, 121, 8)
'     '
'     DBGrid1.COL = 8                     ' OPERARIOS
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     If valor < 0.01 Then valor = 1
'     VALO$ = FORMATNUM$(valor, "F8.2")
'     Call REPLA(TTXX$, VALO$, 129, 8)
'     '
'     DBGrid1.COL = 9                     ' MAQUINA AUX 1
'     CMAQUI$ = TRIM$(DBGrid1.TEXT)
'     If CMAQUI$ <> "" Then
'       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.1'" + CMAQUI$ + "' Inexistente o No Valido"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'     End If
'     Call REPLA(TTXX$, CMAQUI$, 139, 6)
'     '
'     DBGrid1.COL = 10                     ' MAQUINA AUX 2
'     CMAQUI$ = TRIM$(DBGrid1.TEXT)
'     If CMAQUI$ <> "" Then
'       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.2'" + CMAQUI$ + "' Inexistente o No Valido"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'     End If
'     Call REPLA(TTXX$, CMAQUI$, 147, 6)
'     '
'     DBGrid1.COL = 11                     ' MAQUINA AUX 3
'     CMAQUI$ = TRIM$(DBGrid1.TEXT)
'     If CMAQUI$ <> "" Then
'       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.3'" + CMAQUI$ + "' Inexistente o No Valido"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'     End If
'     Call REPLA(TTXX$, CMAQUI$, 155, 6)
'     '
'     DBGrid1.COL = 12                     ' H.P.FABR.
'     HOPROFA$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, HOPROFA$, 163, 8)
'     '
'     List4.AddItem TTXX$
'     Label6 = TRIM$(Str$(List4.ListCount))
'     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'       TEPRUE = RTrim(TTXX$ + "*")
'     End If
'     On Error Resume Next
'     List4.TopIndex = List4.ListCount - 7
'     On Error GoTo 0
'     DoEvents
'     '
'19 Next U&
'   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORMAQ()
   '
'   If DBGrid1.Columns.Count < 7 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   ATX% = 0
'   URECORD& = Data1.Recordset.RecordCount
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(254)
'     DBGrid1.Row = U& - 1
'     DBGrid1.COL = 0
'     CODEX$ = TRIM$(DBGrid1.TEXT)
'     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 6 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     RFF$ = RECFILT$("PADMAQ", 1, CODEX$)
'     If Len(RFF$) > 0 Then
'       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
'       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
'       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
'       If ATX% <= 2 Then GoTo 19 ' SALTEAR
'     End If
'     '
'12   Call REPLA(TTXX$, CODEX$, 1, 6)
'     '
'     DBGrid1.COL = 1
'     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
'     If Len(DEDEX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, DEDEX$, 9, 30)
'     '
'     DBGrid1.COL = 2                     ' Cuota Horaria
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     If valor < 0 Or valor > 999999.99 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     VALO$ = FORMATNUM$(valor, "F12.4")
'     Call REPLA(TTXX$, VALO$, 41, 12)
'     '
'     DBGrid1.COL = 3                    ' Seccion
'     VTX$ = TRIM$(DBGrid1.TEXT)
'     If VTX$ <> "" And XVALO(VTX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion debe ser numerico 'nnn.nnn'"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(VTX$)
'     On Error GoTo 0
'     If valor > 999.999 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion no Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     VALO$ = FORMATNUM$(valor, "F10.3")
'     Call REPLA(TTXX$, VALO$, 53, 10)
'     '
'     DBGrid1.COL = 4                    ' Maquina equivalente
'     CODEQ$ = TRIM$(DBGrid1.TEXT)
'     If CODEQ$ <> "" Then
'       If Len(CODEQ$) < 3 Or Len(CODEQ$) > 6 Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Equivalente No Valido"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'     End If
'     Call REPLA(TTXX$, CODEQ$, 66, 6)
'     '
'     DBGrid1.COL = 5                     ' Coeficiente de equivalencia
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     VAMIN = 0: If CODEQ$ <> "" Then VAMIN = 10
'     If CODEQ$ = "" Then valor = 0
'     If valor < VAMIN Or valor > 999.9 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Coeficiente Equivalencia fuera de Rango"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     VALO$ = FORMATNUM$(valor, "F12.1")
'     Call REPLA(TTXX$, VALO$, 72, 12)
'     '
'     DBGrid1.COL = 6                    ' Observaciones
'     OBSER$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, OBSER$, 87, 30)
'     '
'     List4.AddItem TTXX$
'     Label6 = TRIM$(Str$(List4.ListCount))
'     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'       TEPRUE = RTrim(TTXX$ + "*")
'     End If
'     On Error Resume Next
'     List4.TopIndex = List4.ListCount - 7
'     On Error GoTo 0
'     DoEvents
'     '
'19 Next U&
   
20 End Sub

Sub PREIMPORCAT()
   '
'   If DBGrid1.Columns.Count < 3 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   ATX% = 0
'   URECORD& = Data1.Recordset.RecordCount
'   If URECORD& > 255 Then
'     Call MENSERR(24, "Imposible Importar - Máximo 255 Registros.")
'     GoTo 20
'   End If
'   '
'   If ULTREG&("CATOPER") > 0 Then
'     If COMALTER%("Tabla de Categorias Existente\\Conservar\Sobre-Escribir") <> 2 Then
'       GoTo 20
'     End If
'   End If
'   '
'   jj& = 0
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(254)
'     '
'     DBGrid1.Row = U& - 1
'     DBGrid1.COL = 0
'     DECAT$ = TRIM$(DBGrid1.TEXT)
'     If DECAT$ <> "" Then
'       jj& = jj& + 1
'       COCAT$ = AJUSTD$(Str$(jj&), 4)
'       Call REPLA(TTXX$, COCAT$, 1, 4)
'       Call REPLA(TTXX$, DECAT$, 7, 30)
'       '
'       DBGrid1.COL = 1
'       valor = 0
'       On Error Resume Next
'       valor = CDbl(DBGrid1.TEXT)
'       On Error GoTo 0
'       If valor < 0 Or valor > 999999.99 Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'       VALO$ = FORMATNUM$(valor, "F12.4")
'       Call REPLA(TTXX$, VALO$, 37, 12)
'       '
'       DBGrid1.COL = 2
'       OBSER$ = TRIM$(DBGrid1.TEXT)
'       Call REPLA(TTXX$, OBSER$, 51, 32)
'       '
'       List4.AddItem TTXX$
'       Label6 = TRIM$(Str$(List4.ListCount))
'       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'         TEPRUE = RTrim(TTXX$ + "*")
'       End If
'       On Error Resume Next
'       List4.TopIndex = List4.ListCount - 7
'       On Error GoTo 0
'       DoEvents
'       '
'     End If
'19 Next U&
   
20 End Sub

Sub PREIMPORCLI()
   '
'   If DBGrid1.Columns.Count < 11 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   HROW = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
'   DBGrid1.RowHeight = 10
'   '
'   Screen.MousePointer = 11
'   URECORD& = Data1.Recordset.RecordCount
'   ATX% = 0
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(284)
'     DBGrid1.Row = U& - 1
'     '
'     DBGrid1.COL = 0
'     NUCLIE = XVALO(TRIM$(DBGrid1.TEXT))
'     If NUCLIE < 1 Or NUCLIE > 32000 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cliente no Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     CODEX$ = AJUSTD$(Str$(Int(NUCLIE + 0.5)), 5)
'12   Call REPLA(TTXX$, CODEX$, 1, 5)
'     '
'     DBGrid1.COL = 1
'     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
'     If Len(DEDEX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social de Cliente"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, DEDEX$, 7, 64)
'     '
'     DBGrid1.COL = 2                     ' DOMICILIO
'     DOMI$ = AJUSTI$(DBGrid1.TEXT, 48)
'     Call REPLA(TTXX$, DOMI$, 55, 48)
'     '
'     DBGrid1.COL = 3                     ' CODIGO POSTAL
'     CODPOS$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CODPOS$, 103, 12)
'     '
'     DBGrid1.COL = 4                     ' LOCALIDAD
'     LOCA$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, LOCA$, 115, 32)
'     '
'     DBGrid1.COL = 5                     ' TELEFONO
'     TELEF$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, TELEF$, 147, 24)
'     '
'     DBGrid1.COL = 6                     ' FAX
'     FAX$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, FAX$, 171, 24)
'     '
'     DBGrid1.COL = 7                     ' PERSONA DE CONTACTO
'     CONTACTO$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CONTACTO$, 195, 24)
'     '
'     DBGrid1.COL = 8                     ' CUIT
'     Cuit$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, Cuit$, 219, 16)
'     '
'     DBGrid1.COL = 9                     ' GRUPO DE CLIENTE
'     GRUPOCLI$ = AJUSTI$(DBGrid1.TEXT, 4)
'     Call REPLA(TTXX$, GRUPOCLI$, 236, 4)
'     '
'     DBGrid1.COL = 10                    ' CODIGO DE CUENTA
'     CODICUEN$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CODICUEN$, 241, 12)
'     '
'     DBGrid1.COL = 11                    ' VENDEDOR
'     VENDE$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
'     Call REPLA(TTXX$, VENDE$, 253, 4)
'     '
'     DBGrid1.COL = 12                    ' PROVINCIA
'     PROVI$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, PROVI$, 259, 2)
'     '
'     DBGrid1.COL = 13                    ' CATEGORIA IVA
'     CIVA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
'     Call REPLA(TTXX$, CIVA$, 261, 4)
'     '
'     DBGrid1.COL = 14                    ' CONDICION DE PAGO
'     COPA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
'     Call REPLA(TTXX$, COPA$, 265, 4)
'     '
'     DBGrid1.COL = 15                    ' LIMITE DE CREDITO
'     valor = 0
'     On Error Resume Next
'     valor = CDbl(DBGrid1.TEXT)
'     On Error GoTo 0
'     If valor < 0 Then valor = 0
'     VALO$ = FORMATNUM$(valor, "F12.2")
'     Call REPLA(TTXX$, VALO$, 269, 12)
'     '
'17   List4.AddItem TTXX$
'     Label6 = TRIM$(Str$(List4.ListCount))
'     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'       TEPRUE = RTrim(TTXX$ + "*")
'     End If
'     On Error Resume Next
'     List4.TopIndex = List4.ListCount - 7
'     On Error GoTo 0
'     DoEvents
'19 Next U&
'
'20 Screen.MousePointer = 1
'   '
'   DBGrid1.RowHeight = HROW   ' RESTAURA ALTURA ANTERIOR
   '
End Sub

Sub PREIMPORPRO()
   '
'   If DBGrid1.Columns.Count < 15 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   Screen.MousePointer = 11
'   URECORD& = Data1.Recordset.RecordCount
'   ATX% = 0
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(290)
'     DBGrid1.Row = U& - 1
'     '
'     DBGrid1.COL = 0
'     NUPROV = XVALO(TRIM$(DBGrid1.TEXT))   ' NUMERO DE CUENTA
'     If NUPROV < 1 Or NUPROV > 32000 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor no Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     CODEX$ = AJUSTD$(Str$(Int(NUPROV + 0.5)), 5)
'12   Call REPLA(TTXX$, CODEX$, 1, 5)
'     '
'     DBGrid1.COL = 1                     ' RAZON SOCIAL
'     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
'     If Len(DEDEX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social del Proveedor"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, DEDEX$, 7, 48)
'     '
'     DBGrid1.COL = 2                     ' DOMICILIO
'     DOMI$ = AJUSTI$(DBGrid1.TEXT, 64)
'     Call REPLA(TTXX$, DOMI$, 55, 64)
'     '
'     DBGrid1.COL = 3                     ' CODIGO POSTAL
'     CODPOS$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CODPOS$, 103, 16)
'     '
'     DBGrid1.COL = 4                     ' LOCALIDAD
'     LOCA$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, LOCA$, 115, 48)
'     '
'     DBGrid1.COL = 5                     ' TELEFONO
'     TELEF$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, TELEF$, 147, 24)
'     '
'     DBGrid1.COL = 6                     ' FAX
'     FAX$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, FAX$, 171, 24)
'     '
'     DBGrid1.COL = 7                     ' PERSONA DE CONTACTO
'     CONTACTO$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CONTACTO$, 195, 24)
'     '
'     DBGrid1.COL = 8                     ' CUIT
'     Cuit$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, Cuit$, 219, 16)
'     '
'     DBGrid1.COL = 9                     ' TIPO DE CUENTA
'     TIPOCTA$ = TRIM$(DBGrid1.TEXT)
'     If TIPOCTA$ = "0" Or TIPOCTA$ = "1" Then
'       Call REPLA(TTXX$, TIPOCTA$, 235, 4)
'     Else
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Tipo de Cuenta"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     '
'     DBGrid1.COL = 10                    ' CODIGO DE CUENTA
'     CODICUEN$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CODICUEN$, 240, 12)
'     '
'     '\\\OT-05-0202-WAR-18/04/05///
'     DBGrid1.COL = 11                    ' CODIGO DE PROVINCIA
'     CODIPROVI$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, CODIPROVI$, 253, 2)
'     '
'     DBGrid1.COL = 12                    ' CODIGO DE POSICION IVA
'     CODIVA% = XVALO(DBGrid1.TEXT)
'     Call REPLA(TTXX$, FORMATNUM$(CODIVA%, "I4"), 255, 4)
'     '\\\OT-05-0202-WAR-FIN///
'     '
'     DBGrid1.COL = 13                    ' CUENTA MADRE
'     CTAMADRE$ = TRIM$(DBGrid1.TEXT)
'     If Len(CTAMADRE$) > 12 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Longitud de Cuenta Madre"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, CTAMADRE, 261, 12)
'     '\\\OT-05-0202-WAR-FIN///
'     '
'     DBGrid1.COL = 14                    ' numbre de fantasia
'     FANTASI$ = TRIM$(DBGrid1.TEXT)
'     Call REPLA(TTXX$, FANTASI$, 275, 16)
'     '
'    '
'17   List4.AddItem TTXX$
'     Label6 = TRIM$(Str$(List4.ListCount))
'     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
'       TEPRUE = RTrim$(TTXX$ + "*")
'     End If
'     On Error Resume Next
'     List4.TopIndex = List4.ListCount - 7
'     On Error GoTo 0
'     DoEvents
'19 Next U&
'
'20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORPLA()
   '
'   If DBGrid1.Columns.Count < 5 Then
'     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
'     GoTo 20
'   End If
'   '
'   HROW = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
'   DBGrid1.RowHeight = 10
'   '
'   ATX% = 0
'   URECORD& = Data1.Recordset.RecordCount
'   For U& = 1 To URECORD&
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(254)
'     DBGrid1.Row = U& - 1
'     DBGrid1.COL = 0
'     '
'     NIVEI% = XVALO(DBGrid1.TEXT)
'     If NIVEI% < 0 Or NIVEI% > 5 Then NIVEI% = 0
'     Call REPLA(TTXX$, AJUSTD$(Str$(NIVEI%), 4), 1, 4)
'     '
'     DBGrid1.COL = 1
'     CODEX$ = TRIM$(DBGrid1.TEXT)
'     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 12 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Inexistente o No Valido"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, CODEX$, 7, 12)
'     '
'     For KKU% = 1 To List4.ListCount
'       If TRIM$(Mid$(List4.LIST(KKU% - 1), 7, 12)) = CODEX$ Then
'         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Duplicado o Pre-Existente"
'         List1.AddItem MERROX$
'         GoTo 19
'       End If
'     Next KKU%
'     '
'     DBGrid1.COL = 2
'     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
'     If Len(DEDEX$) < 1 Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
'     Call REPLA(TTXX$, DEDEX$, 21, 48)
'     '
'     DBGrid1.COL = 3                    ' Seccion
'     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
'     If VTX$ <> "D" And VTX$ <> "H" And VTX$ <> "I" Then VTX$ = "G"
'     Call REPLA(TTXX$, VTX$, 71, 1)
'     '
'     DBGrid1.COL = 4
'     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
'     If VTX$ = "0" Or VTX$ = "NO" Then VTX$ = "NO"
'     If VTX$ <> "NO" Then VTX$ = "SI"
'     Call REPLA(TTXX$, VTX$, 75, 2)
'     '
'     List4.AddItem TTXX$
'     Label6 = TRIM$(Str$(List4.ListCount))
'     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
'       TEPRUE = RTrim(TTXX$ + "*")
'     End If
'     On Error Resume Next
'     List4.TopIndex = List4.ListCount - 7
'     On Error GoTo 0
'     DoEvents
'     '
'19 Next U&
'   '
'   DBGrid1.RowHeight = HROW   ' RESTAURA ALTURA ANTERIOR
   '
20 End Sub
 '\\\OT-05-0207-WAR-21/04/05///
Sub PREIMPOPED()
   '
   
   If DBGrid1.Cols < 6 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DBGrid1.Rows
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     '
     DBGrid1.COL = 3            'CANTIDAD
     CANTU# = 0
     On Error Resume Next
     CANTU# = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If CANTU# < 0.05 Then GoTo 19
     '
     DBGrid1.COL = 0            'CODIGO ARTI
     CODEX$ = TRIM$(DBGrid1.TEXT)
     Ci% = CODINT(CODEX$)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or Ci% <= 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo '" + CODEX$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     '
     DBGrid1.COL = 1                     ' NUMERO ORIGINAL
     VALO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, VALO$, 18, 28)
     '
     DBGrid1.COL = 2            'DESCRI ARTI
     DEDEX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DEDEX$, 48, 64)
     '
     DBGrid1.COL = 3            'CANTIDAD
     CANTU# = 0
     On Error Resume Next
     CANTU# = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If CANTU# < 0.05 Then GoTo 19
     CANTI1$ = FORMATNUM(CANTU#, "F12.1")
     Call REPLA(TTXX$, CANTI1$, 114, 12)
     '
     DBGrid1.COL = 4            'FECHA DE ENTREGA
     FECHENT$ = FECHATEX$(CVINT(DBGrid1.TEXT))
     If TRIM(FECHENT$) = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Fecha de Entrega no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, FECHENT$, 128, 8)
     '
     DBGrid1.COL = 5            'PRECIO UNITARIO
     PUNI# = 0
     On Error Resume Next
     PUNI# = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If PUNI# <= 0 Then
        CI1% = CODINT(CODEX$)
        ARCHIPRE$ = "LIPRE"
        L1% = XVALO(Label11)
        L2$ = TRIM$(FORMATNUM$(L1%, "I5"))
        Do While Len(L2$) < 3
          L2$ = "0" + L2$
        Loop
        ARCHIPRE1$ = ARCHIPRE$ + L2$
        PUNI# = PRELISTA(ARCHIPRE1$, CI1%)
     End If
     PUNI1$ = FORMATNUM(PUNI#, "F14.4")
     Call REPLA(TTXX$, PUNI1$, 138, 14)
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
Sub IMPORPED()
   Dim FORMORIG As Form
   Set FORMORIG = FORPED07
   If PROLLAMA = "ABIE" Then Set FORMORIG = PEDABIE07
    With FORMORIG
        Call .CARDACLI
  'CARGO EL DETALLE
  Screen.MousePointer = 11
  '
  REBLA$ = REGBLAN$("!CARDETPE")
  J& = 0
  For KI% = 1 To List4.ListCount
    TTXX$ = List4.List(KI% - 1)
    NCI2$ = (Mid$(TTXX$, 1, 16)) 'Nº ARTICULO
    NCI1% = CODINT(NCI2$)
    JJD$ = Mid$(TTXX$, 114, 12)
    CA1# = XVALO(Mid$(TTXX$, 114, 12))  'CANTIDAD
    FE% = FECHANUM(Mid$(TTXX$, 128, 8)) 'FECHA
    PU1# = XVALO(Mid$(TTXX$, 136, 14)) 'PRECIO UNITARIO
    Call REPLA(Z$, MKI(NCI1%), 1, 2)
    Call REPLA(Z$, MKS$(CA1#), 47, 4)
    Call REPLA(Z$, MKI$(FE%), 63, 2)
    Call REPLA(Z$, MKS$(PU1#), 39, 4)
    Call REPLA(Z$, MKS$(PU1#), 73, 4)
    J& = J& + 1
    Call GRAREG("!CARDETPE", Z$, J&)
  Next KI%
  Call SETULTREG("!CARDETPE", J&)
  '
  On Error Resume Next
  Call .CARLISPE
  .MARCONOTA.Visible = False
  .MARSELPRESU.Visible = False
  Call .CALTOPED
  Call .CALPERIB
  End With
  On Error GoTo 0
  IMPPED07.Hide
  Screen.MousePointer = 1
  '
End Sub
