VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IMPSTO07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de ImportaciÓn de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   11040
      ScaleHeight     =   165
      ScaleWidth      =   720
      TabIndex        =   44
      Top             =   2280
      Width           =   750
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   43
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "ORIGEN DE DATOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   2630
      Width           =   11655
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
         Left            =   8040
         TabIndex        =   37
         Top             =   480
         Width           =   220
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         ScaleHeight     =   285
         ScaleWidth      =   7785
         TabIndex        =   35
         Top             =   480
         Width           =   7815
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   15
            TabIndex        =   36
            Top             =   0
            Width           =   15000
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   10970
         Picture         =   "IMPSTO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Establece Conexion con Origen de Datos"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   33
         Text            =   " "
         Top             =   460
         Width           =   2295
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8270
         TabIndex        =   38
         Top             =   240
         Width           =   735
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
      BackColor       =   &H00CDDADA&
      Caption         =   "RESULTADO DE PRE-IMPORTACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Importación Definitiva"
      Top             =   3720
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   20
         Top             =   420
         Width           =   10215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -20
            ScaleHeight     =   1965
            ScaleWidth      =   10215
            TabIndex        =   22
            Top             =   0
            Width           =   10215
            Begin VB.PictureBox Picture3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1950
               Left            =   -20
               ScaleHeight     =   1920
               ScaleWidth      =   30
               TabIndex        =   27
               Top             =   0
               Width           =   60
            End
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   0
               ScaleHeight     =   240
               ScaleWidth      =   1.19970e5
               TabIndex        =   24
               Top             =   0
               Width           =   1.20000e5
               Begin VB.Label Label83 
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
                  Left            =   150
                  TabIndex        =   28
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
                  Left            =   40
                  TabIndex        =   26
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
                  TabIndex        =   25
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.ListBox List4 
               Appearance      =   0  'Flat
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
               ItemData        =   "IMPSTO07.frx":030A
               Left            =   0
               List            =   "IMPSTO07.frx":030C
               TabIndex        =   23
               Top             =   250
               Width           =   1.30000e5
            End
         End
         Begin VB.PictureBox Picture10 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   21
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         IntegralHeight  =   0   'False
         ItemData        =   "IMPSTO07.frx":030E
         Left            =   120
         List            =   "IMPSTO07.frx":0310
         TabIndex        =   17
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
         Picture         =   "IMPSTO07.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Listado por Impresora"
         Top             =   3570
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
         TabIndex        =   15
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
         Left            =   10710
         Picture         =   "IMPSTO07.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1785
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1725
         Left            =   10360
         TabIndex        =   9
         Top             =   670
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   8
         Top             =   2340
         Width           =   10250
      End
      Begin VB.Label label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
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
         Left            =   6000
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTROS NO PROCESABLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2670
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   13
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10690
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "Planilla Origen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   6
      Top             =   3675
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid DbGrid1 
         Height          =   4100
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMATO DE IMPORTACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   6100
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         ScaleHeight     =   270
         ScaleWidth      =   5625
         TabIndex        =   4
         Top             =   350
         Width           =   5655
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   5
            Top             =   0
            Width           =   5655
         End
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
      Left            =   11025
      Picture         =   "IMPSTO07.frx":0926
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
      Height          =   330
      Left            =   1785
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   0
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
      Picture         =   "IMPSTO07.frx":0C30
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   19
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "IMPSTO07.frx":0D7A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "IMPSTO07.frx":0FAE
      TabIndex        =   42
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1695
      Left            =   1680
      TabIndex        =   31
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "de Inventarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   645
      Left            =   1050
      TabIndex        =   30
      Top             =   1050
      Width           =   4005
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Importador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   645
      Left            =   315
      TabIndex        =   29
      Top             =   420
      Width           =   4005
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPSTO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CANCELIMP%
'
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private HOJA As Excel.Worksheet
'
Sub PREIMPINVSTO()
   '
   ATX% = 0
   '
   U& = 1
   '  Do While Not .Recordset.EOF
      Do While U& < URECORD&
        U& = U& + 1
        DbGrid1.row = U& - 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(254)
        CODEX$ = ""
        On Error Resume Next
        DbGrid1.COL = 0
        CODEX$ = TRIM$(DbGrid1.TEXT) ' .Recordset.Fields(0))   ' TRIM$(DBGrid1.TEXT)
     Loop
   '03/04/07 (OT-07-0141)
   
   If DbGrid1.Cols < 5 Then
      Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
      GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
'   HROW = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
'   DbGrid1.RowHeight = 10
   '
   ATX% = 0
'   HUBOAMAS% = 0
   '
'   With Data1
'     .Recordset.MoveFirst
     U& = 1
     Do While U& < URECORD
        U& = U& + 1
        DbGrid1.row = U& - 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(120)
        '
        CODEX$ = ""                             ' CODIGO DE ARTICULO
        On Error Resume Next
        CODEX$ = DbGrid1.TextMatrix(U& - 1, 0) '.Recordset.Fields(0)
        On Error GoTo 0
        '
5       CDXI% = CODINT%(CODEX$)
        If CDXI% < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Código '" + TRIM$(CODEX$) + "' No Existente en el Maestro de Artículos"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
12      CODEX$ = CODEXT$(CDXI%)
        Call REPLA(TTXX$, CODEX$, 1, 16) 'CODIGO DEL ARTICULO
        '
        Call REPLA(TTXX$, DESCRIT0$(CDXI%), 19, 48) ' DESCRIPCION DEL ARTICULO
        '
        FECHMOVX$ = ""
        On Error Resume Next
        fech% = CVINT(DbGrid1.TextMatrix(U& - 1, 2)) ' .Recordset.Fields(2))
        FECHMOVX$ = FECHATEX$(fech%) 'FECHA
        On Error GoTo 0
        If FECHMOVX$ = "" Or FECHANUM(FECHMOVX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Fecha No Válida"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, FECHMOVX$, 67, 8)
        '

        DEPX% = 1 ' DEPóSITO
        On Error Resume Next
        DEPX% = CDbl(DbGrid1.TextMatrix(U& - 1, 3)) ' .Recordset.Fields(3))
        On Error GoTo 0
        If DEPX% < 1 Or DEPX% > 255 Then DEPX% = 1
        VALO$ = FORMATNUM$(DEPX%, "F4.0")
        Call REPLA(TTXX$, VALO$, 75, 4)
        '
        ' *** VALIDA LA EXISTENCIA DEL LOTE EN ARTICULOS TRAZABLES
        ' *** PERMITE LOTES VACIOS PARA ARTICULOS TRAZABLES, YA QUE EL IMPORTADOR SE ENCARGA DE
        ' *** CREAR UN MOVIMIENTO DE ENTRADA DONDE EL LOTE ES EL PROPIO NUMERO DE AJUSTE
        If ESTRAZABLE%(CDXI%) > 0 Then
            valor = 0
            On Error Resume Next
            valor = DbGrid1.TextMatrix(U& - 1, 5) ' .Recordset.Fields(5)
            On Error GoTo 0
            LOTE$ = TRIM$(valor)
            IExisteLote% = CantRegistros("LOTINGRE", "IDENLOTE='" & LOTE$ & "' AND COD_INTE=" & CStr(CDXI%))
            If (LOTE$ <> "" And IExisteLote% = 0) Then
              MERROX$ = "Registro" + Str$(U& - 1) + " - Lote No Válido"
              List1.AddItem MERROX$
              GoTo 19
            End If
            Call REPLA(TTXX$, AJUSTI$(LOTE$, 16), 105, 16)
        End If
        'FIN LOTE

        valor = 0 'STOCK TEORICO
        valor = STODEPFECH(CDXI%, DEPX%, fech%, LOTE$) '04/04/07 (OT-07-0141)
        VALO$ = FORMATNUM$(valor, "F12.1")
        Call REPLA(TTXX$, VALO$, 79, 12)
        '
        valor = 0 'STOCK AJUSTADO
        On Error Resume Next
        valor = CDbl(DbGrid1.TextMatrix(U& - 1, 4)) ' .Recordset.Fields(4))
        On Error GoTo 0
        VALO$ = FORMATNUM$(valor, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Ajustado No Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 91, 12)
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
19  ' .Recordset.MoveNext
     Loop
   ' End With
   '
20 ' DbGrid1.RowHeight = HROW
   Screen.MousePointer = 1
   '
End Sub
'
Private Sub Combo1_Change()
'   '
'   Data1.RecordSource = Combo1.TEXT
'   Data1.Refresh
'   Data1.Recordset.MoveFirst
'   '
End Sub

Private Sub Combo1_Click()
   '
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.List(KKI% - 1) = Libro.Sheets(KKI%).Name Then
        Set HOJA = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set HOJA = Libro.Sheets(1)
   DbGrid1.Clear
   DbGrid1.Rows = 1
   '
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
   CANCELIMP% = 0
   '
   Call IMPINVSTO
   '
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
   URECORD& = DbGrid1.Rows   '   Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DbGrid1.Height
   '
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   '
   Call PREIMPINVSTO
   '
   Call COMUNI("Se Ha Completado la Pre-Importación.")
   '
20 DbGrid1.Height = DBHANT
   Command4.Enabled = True
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.max = MXX%
   '
   HScroll1.Value = 0
   HScroll1.max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.max = (TEPRUE.Width - HScroll1.Width) / 10
   End If
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
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
   For i& = 1 To CANFIL&
     YAGRELI% = 0
     For J& = 1 To CANCOL&
       On Error Resume Next
       valor = HOJA.Cells(i&, J&)
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Imposible Establecer Origen de Datos.")
          Screen.MousePointer = 1
          Exit Sub
       End If
       On Error GoTo 0
       '
       If TRIM$(CStr(valor)) <> "" Then
         If YAGRELI% < 1 Then
           YAGRELI% = 1
           JK& = JK& + 1
         End If
         If JK& > DbGrid1.Rows Then DbGrid1.Rows = DbGrid1.Rows + 1
         DbGrid1.row = JK& - 1
       End If
         If J& > DbGrid1.Cols Then DbGrid1.Cols = DbGrid1.Cols + 1
       DbGrid1.COL = J& - 1
       If TRIM$(valor) <> "" Then
          DbGrid1.TextMatrix(JK& - 1, J& - 1) = valor
       End If
       '
       Label9.Font = DbGrid1.Font
       Label9.FontSize = DbGrid1.CellFontSize
       Label9.FontBold = DbGrid1.CellFontBold
       Label9.Caption = CStr(valor) + "**"
       If Label9.Width > DbGrid1.ColWidth(J& - 1) Then
         DbGrid1.ColWidth(J& - 1) = Label9.Width
       End If
     Next J&
     If i& > JK& + 256 Then
        GoTo 99
     End If
   Next i&
   '
99 Screen.MousePointer = 1
   Call COMUNI("Origen de Datos Establecido.")
   '
End Sub

Private Sub Form_Load()
   '03/04/07 (OT-07-0141)
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)

   Call CENTRAFORM(Me)
   '
   List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
   List2.AddItem "  2  Descripcion del Articulo   String   <= 64 car"
   List2.AddItem "  3  Fecha                      Date"
   List2.AddItem "  4  Depósito Número            Num      1 - 255"
   List2.AddItem "  5  Stock Ajustado             Num      +/- 999999"
   List2.AddItem "  6  Lote                       String   <= 16 car"
   Label8 = "Código            Descripción                                     Fecha     Dep     S.Teo.     S.Ajus.  Lote"
   '03/04/07 (FIN)
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

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
End Sub




Private Sub label1_Change()
  DbGrid1.Clear
  DbGrid1.Rows = 1
  Combo1.Clear
  Combo1.TEXT = ""
  '
  If Label1.Caption <> "" Then
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
  End If
  
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

Sub IMPINVSTO()
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT DoPriCor FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiMovi = 'AJ' "
   'SQLX$ = SQLX$ + "ORDER BY VAL(SUBSTRING(DoPriCor,4,)) DESC"
   SQLX$ = SQLX$ + " ORDER BY CAST(SUBSTRING(DoPriCor,4,12) AS INTEGER) DESC" ' SE ACTUALIZO PARA SQL

   UNUM& = 0
   'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim MoviTemp As ADODB.Recordset
'   Set MoviTemp = New ADODB.Recordset
'   MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set MoviTemp = READSET(SQLX$)
   On Error Resume Next
   MoviTemp.MoveFirst
   If Err < 1 Then
      UNUM& = XVALO(Mid$(MoviTemp!dopricor, 4))
   End If
   On Error GoTo 0
   NRO& = 1 + UNUM&
   MoviTemp.Close
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NOCLOSE% = 1
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     CIXI% = CODINT%(CODIT$)
     FEMO% = FECHANUM(Mid$(TTXX$, 67, 8))
     CMOV$ = "AJ"
     DOSECO$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
     DOPRICO$ = "AS." + TRIM$(Str$(NRO&))
     DOPRILA$ = DOPRICO$
     DOSELA$ = DOSECO$
     REFE$ = "Importación Inventario al " + FECHATEX$(HOY(HO$))
     NUORIT% = 0
     VUNI# = 0
     MONEII% = 0
     NC% = 0
     depo% = XVALO(Mid$(TTXX$, 75, 4))
     CANTI = XVALO(Mid$(TTXX$, 91, 12)) - XVALO(Mid$(TTXX$, 79, 12))
     LOTE$ = TRIM$(Mid$(TTXX$, 105, 16))
     
     If Abs(CANTI) >= 0.005 Then
        '
        ' SI NO TIENE LOTE Y ES TRAZABLE, EL AJUSTE DE ENTRADA TIENE COMO LOTE A SI MISMO
        If ESTRAZABLE%(CIXI%) > 0 And TRIM$(LOTE$) = "" Then LOTE$ = DOPRILA$
        '
        Call MOVISTOK(FEMO%, CIXI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, depo%, CANTI)
        '
        ' AGREGA EL NUEVO LOTE A LA TABLA LOTINGRE
        If ESTRAZABLE%(CIXI%) > 0 And LOTE$ = DOPRILA$ Then Call CREALOTIN(FEMO%, 0, 0, CIXI%, LOTE$, "", "", CDbl(CANTI))
        
        NRO& = NRO& + 1
     End If
     '
   Next U&
   '
   'Movisto.Close
   NOCLOSE% = 0
   Screen.MousePointer = 1
   '
End Sub

