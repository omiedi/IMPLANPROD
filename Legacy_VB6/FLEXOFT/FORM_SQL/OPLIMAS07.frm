VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OPLIMAS07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Items de Stock - Consultas, Listados"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command18 
      Caption         =   "Baja x Excel"
      Height          =   255
      Left            =   3550
      TabIndex        =   39
      ToolTipText     =   "Da de Baja Items Importados desde una Archivo Excel"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   255
      Left            =   3960
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   36
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command8 
      Caption         =   "IMP."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5080
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   4935
      ScaleHeight     =   7545
      ScaleWidth      =   1035
      TabIndex        =   24
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   40
         Top             =   6360
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPLIMAS07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Cálculo de Stock Minimo por Autonomía Proyectada y Consumo Promedio"
         Top             =   2400
         Width           =   640
      End
      Begin VB.CommandButton Command14 
         Caption         =   "D.Art."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Proceso Para Determinar Estado de Artículos"
         Top             =   4875
         Width           =   645
      End
      Begin VB.CommandButton Command13 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Importar Consumo Promedio\Importación de UPC"
         Top             =   4425
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Excel 5.0;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Imp. M2/U"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Importa Metros Cuadrados por Unidad (Etiquetas)"
         Top             =   3195
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "OPLIMAS07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   5325
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Refr."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPLIMAS07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Refresca Lista de Seleccion de Articulos"
         Top             =   1590
         Width           =   640
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   28
         ToolTipText     =   "Análisis de Costos de Reposición"
         Top             =   3645
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPLIMAS07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   300
         Width           =   640
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPLIMAS07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   940
         Width           =   640
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "ITEMS INACTIVOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   4635
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   615
         ScaleWidth      =   2610
         TabIndex        =   9
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "OPLIMAS07.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Listado de Items Inactivos (Dados de Baja)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Items Inactivos (de Baja)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   11
            Top             =   90
            Width           =   1785
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2610
         TabIndex        =   12
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "OPLIMAS07.frx":0EBC
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Listado de Items Inactivos (Dados de Baja)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta y Restauración Items de Baja"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   13
            Top             =   60
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         Caption         =   "COSTO DE REPOSICION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   14
         Top             =   2730
         Width           =   4215
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   615
            ScaleWidth      =   2610
            TabIndex        =   15
            Top             =   945
            Width           =   2640
            Begin VB.CommandButton Command4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLIMAS07.frx":1006
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Listado de Costos de Reposición por Articulo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado por  Código de Artículo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   750
               TabIndex        =   17
               Top             =   100
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   630
            ScaleHeight     =   615
            ScaleWidth      =   2610
            TabIndex        =   18
            Top             =   420
            Width           =   2640
            Begin VB.CommandButton Command5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLIMAS07.frx":1310
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Listado de Costos de Reposición por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado por Número de Proveedor"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   20
               Top             =   60
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "LISTADOS GENERALES DEL MAESTRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   4215
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   615
            ScaleWidth      =   2610
            TabIndex        =   5
            Top             =   1470
            Width           =   2640
            Begin VB.CommandButton Command15 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLIMAS07.frx":161A
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Captura de Datos para Inventarios"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Captura para Inventarios"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   700
               TabIndex        =   7
               Top             =   100
               Width           =   1905
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2610
            TabIndex        =   21
            Top             =   945
            Width           =   2640
            Begin VB.CommandButton Command7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLIMAS07.frx":1924
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Listado de Atributos Generales de Articulos de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Stock Mínimo y Lote de Reposición"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   23
               Top             =   40
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   630
            ScaleHeight     =   615
            ScaleWidth      =   2610
            TabIndex        =   2
            Top             =   420
            Width           =   2640
            Begin VB.CommandButton Command9 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLIMAS07.frx":1D66
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Listado de Atributos Generales de Articulos de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Atributos Generales"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   4
               Top             =   40
               Width           =   1695
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPLIMAS07.frx":1EB0
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "OPLIMAS07.frx":20E4
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPLIMAS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private HOJA As Excel.Worksheet

Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command10_Click()
    '
    UPDATMASTER% = 1
    Call GENECOMAS
    Call FRESHECHO("ECOMAST")
    Call COMPACTATABLA("TATIMAT")
    Call COMPACTATABLA("UNIMED")
    Call COMPACTATABLA("UMEDIDA")
    Call MENSERR(24, "Proceso Finalizado")
    '
End Sub

Private Sub Command11_Click()
    '
    If InStr(1, EMPREAC$, "ARCON") > 0 Then
        Call ABRECONEXION
        '
        Dim MovistoTemp As ADODB.Recordset
        Set MovistoTemp = New ADODB.Recordset
        '
        FF% = FECHANUM("20/01/12")
        SQLX$ = "SELECT * FROM MOVISTO WHERE CODIMOVI='RT' "
        SQLX$ = SQLX$ + " AND (DOPRICOR='RT.1288' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1289' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1290' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1291' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1292' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1293' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1294' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1295' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1296' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1297' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1298' "
        SQLX$ = SQLX$ + " OR DOPRICOR='RT.1299') "
        SQLX$ = SQLX$ + " AND FECHMOV >'" & FETEXCOR1$(FF%) & "'"
        '
        MovistoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        Do While Not MovistoTemp.EOF
            MovistoTemp!DEPOMOVI = 10
            MovistoTemp.Update
            MovistoTemp.MoveNext
        Loop
        '
        MovistoTemp.Close
        Set MovistoTemp = Nothing
    End If
    '
End Sub

Private Sub Command12_Click()
Command12.Enabled = False
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0

    DoEvents
    Screen.MousePointer = 11
    Data1.DatabaseName = UCase$(CommonDialog1.FileName)
    On Error Resume Next
    Data1.Refresh
    Data1.RecordSource = Data1.Database.TableDefs(0).Name
    Data1.Refresh
    On Error GoTo 0
    
    With Data1
     .Recordset.MoveFirst
     Do While Not .Recordset.EOF
        Codigo$ = ""
        On Error Resume Next
        Codigo$ = .Recordset.Fields(0)
        Ci% = CODINT%(Codigo$)
        If Ci% < 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente.\Dar de Alta y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        On Error GoTo 0
        M2UNI# = .Recordset.Fields(1)
        M2XUNID$ = FORMATNUM$(M2UNI#, "F12.8")
        Call ACTUREGISTRO("MASTER", "Pesuni", "Codint = " & Ci%, M2XUNID$, REGAF&, "NOMESS")
20      Screen.MousePointer = 11
       .Recordset.MoveNext
     Loop
   End With
   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
99 Command12.Enabled = True

End Sub

Private Sub Command13_Click()
    '
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    Dim IMPORTACION As Boolean
    IMPORTACION = False
    URECORD& = UltimoRegistro(1)
    For Fila& = 2 To URECORD&
        '
        On Error Resume Next
        Codigo$ = TRIM$(HOJA.Cells(Fila&, 1))
        Ci% = CODINT%(Codigo$)
        CONSPROM# = XVALO(HOJA.Cells(Fila&, 3))
        On Error GoTo 0
        If Ci% < 1 Then
          If CONSPROM# >= 0.005 Then
            Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente.\Dar de Alta y Volver a Correr el Proceso de Importación")
          End If
          GoTo 20
        End If
        Call ACTUREGISTRO("MASTER", "consuprom", "Codint = " & Ci%, CONSPROM#, REGAF&, "NOMESS")
20 IMPORTACION = True
   Next Fila&
   '
   Libro.Close
   objexcel.Quit
   Set Libro = Nothing
   Set objexcel = Nothing
   '
   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
99 Command13.Enabled = True
End Sub



Function SELECCARCHIVOEXCEL$()
    SELECCARCHIVOEXCEL$ = ""
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    SELECCARCHIVOEXCEL$ = UCase$(CommonDialog1.FileName)
99
End Function

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    Dim IMPORTACION As Boolean
    IMPORTACION = False
    URECORD& = UltimoRegistro(1)
    For Fila& = 2 To URECORD&
        '
        Call TRACE(24, Fila&, URECORD&)
        On Error Resume Next
        Codigo$ = TRIM$(HOJA.Cells(Fila&, 1))
        Ci% = CODINT%(Codigo$)
        On Error GoTo 0
        If Ci% < 1 Then
            Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente.\Dar de Alta y Volver a Correr el Proceso de Importación")
            GoTo 20
        End If
        Call ACTUREGISTRO("MASTER", "Status", "Codint = " & Ci%, (-1), REGAF&, "")
20 IMPORTACION = True
   Next Fila&
   '
   Libro.Close
   objexcel.Quit
   Set Libro = Nothing
   Set objexcel = Nothing
   '
   UPDATMASTER% = 1
   Call GENECOMAS
   Call FRESHECHO("ECOMAST")
   Call COMPACTATABLA("TATIMAT")
   Call COMPACTATABLA("UNIMED")
   Call COMPACTATABLA("UMEDIDA")
   Call MENSERR(24, "Proceso Finalizado")
   '
   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
   
99 Command14.Enabled = True
'    '
'    Command14.Enabled = False
'    '
'    ' SI NO TIENE DERECHOS DE USUARIO NO CORRE EL PROCESO
'    If DERACCE%("DEPUMOVI", "Proceso De Depuración de Artículos Segun Movimientos") < 2 Then
'      GoTo 99
'    End If
'    '
'    Screen.MousePointer = 11
'    '
'    'GENERO EL RANGO DE FECHAS DEL MES ANTERIOR
'    DES% = COMESANT%
'    HAS% = FINMESANT%
'    DESDE = FETEXCOR1$(DES%)
'    HASTA = FETEXCOR1$(HAS%)
'    DESHASTA$ = FECHATEX$(DES%) + " al " + FECHATEX$(HAS%) ' TEXTO DE PERIODO
'    '
'    Call ABREMASTER
'    Call ABRESTOCKS
'    '
'    ' Creo aqui los  3 campos.
'    Call CREACAMPO("ARTICULOS", "MASTER", "PERIOBAJA", 10, 32) 'DESDE-HASTA COMO TEXTO
'    Call CREACAMPO("ARTICULOS", "MASTER", "CAMOVSALI", 7, 0) 'CANTIDAD MOVIMIENTO DE SALIDA
'    Call CREACAMPO("ARTICULOS", "MASTER", "IMOVSALI", 7, 0) ' IMPORTE MOVIMIENTO DE SALIDA
'    '
'    SQLX$ = "SELECT DISTINCT(CODINT) FROM MASTER  INNER JOIN PREVENTA "
'    SQLX$ = SQLX$ + " ON CODINT = COD_PIEZA"
'    SQLX$ = SQLX$ + " WHERE STAT_MOVI <> -1 AND STAT_MOVI <> 1 "
'    Dim MOVTMP As ADODB.Recordset
'    Set MOVTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    With MOVTMP
'      Do While Not .EOF
'        CI% = XVALO(!CODINT)
'        ' CONSULTO EN EL MOVISTO SI TIENE MOVIMIENTOS
'        CONDI$ = " COD_INTE = " & CI%& " AND FECHMOV >=  '" & DESDE & "'"
'        CONDI$ = CONDI$ + "AND FECHMOV <=  '" & HASTA & "'"
'        CANTISALI = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
'        IMPORTE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID * VALOUNIT)", CONDI$ + " AND CODIMOVI = 'RT'", "NOMESS"))
'        If CANTISALI < 5000 Then
'           If IMPORTE < 5000 Then
'                Call ACTUREGISTRO("MASTER", "Stat_Movi", "CODINT = " & CI%, (-1), REAF&, "NOMESS")
'                Call ACTUREGISTRO("MASTER", "PERIOBAJA", "CODINT = " & CI%, DESHASTA$, REAF&, "NOMESS")
'                Call ACTUREGISTRO("MASTER", "CAMOVSALI", "CODINT = " & CI%, CANTISALI, REAF&, "NOMESS")
'                Call ACTUREGISTRO("MASTER", "IMOVSALI", "CODINT = " & CI%, IMPORTE, REAF&, "NOMESS")
'           End If
'        End If
'        .MoveNext
'        Loop
'    End With
'    '
'    MOVTMP.Close
'    Set MOVTMP = Nothing
'    '
'    Call COMUNI("Proceso Terminado")
'    '
'99  Screen.MousePointer = 1
'    Command14.Enabled = True
'    '
End Sub

Private Sub Command16_Click()
    
    Command16.Enabled = False
    '
    STOMPROM.Show 1
    '
99  Command16.Enabled = True

End Sub



Private Sub Command17_Click()
    '
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    Dim IMPORTACION As Boolean
    IMPORTACION = False
    URECORD& = UltimoRegistro(1)
    For Fila& = 2 To URECORD&
        '
        Call TRACE(24, Fila&, URECORD&)
        On Error Resume Next
        Codigo$ = TRIM$(HOJA.Cells(Fila&, 1))
        Ci% = CODINT%(Codigo$)
        LEADTIME# = XVALO(TRIM$(HOJA.Cells(Fila&, 5)))
        If LEADTIME# > 0 Then
        A = B
        End If
        STOMINI# = XVALO(TRIM$(HOJA.Cells(Fila&, 6)))
        On Error GoTo 0
        If Ci% < 1 Then GoTo 20
        
        SQLX$ = "select * from master where codint = " & Ci%
        Dim Rstactucos As ADODB.Recordset
        Set Rstactucos = New ADODB.Recordset
        Rstactucos.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Not (Rstactucos.BOF And Rstactucos.EOF) Then
           Rstactucos!LETICOM = LEADTIME#
           Rstactucos!STOMIN = STOMINI#
           Rstactucos.Update
        End If
        Rstactucos.Close
        Set Rstactucos = Nothing
        
20 IMPORTACION = True
   Next Fila&
   '
   Libro.Close
   objexcel.Quit

   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")

99 Command8.Enabled = False

End Sub

Private Sub Command18_Click()
    '
    If CLACONTRA% < 1 Then Exit Sub
    
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    Dim IMPORTACION As Boolean
    IMPORTACION = False
    URECORD& = UltimoRegistro(1)
    For Fila& = 1 To URECORD&
        '
        Call TRACE(24, Fila&, URECORD&)
        On Error Resume Next
        Codigo$ = TRIM$(HOJA.Cells(Fila&, 1))
        Ci% = CODINT%(Codigo$)
        If Ci% < 1 Then GoTo 20
        CONDI$ = "CODINT = " & Ci%
        Call ACTUREGISTRO("MASTER", "STATUS", CONDI$, (-1), REGAF&, "NOMESS")
20      IMPORTACION = True
   Next Fila&
   '
   Libro.Close
   objexcel.Quit

   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Baja Finalizado")
   Call GENEECOMABAJ 'ACTUALIZA LISTA DE SELECCION
   '
   Call GENECOMAS

99 Command18.Enabled = True

End Sub


Private Sub Command8_Click()
    '
    Set objexcel = New Excel.Application
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    '
    Ruta$ = UCase$(CommonDialog1.FileName)
    Set Libro = objexcel.Workbooks.Open(Ruta$)
    Set HOJA = Libro.Sheets(1)
    '
    Dim IMPORTACION As Boolean
    IMPORTACION = False
    URECORD& = UltimoRegistro(1)
    For Fila& = 1 To URECORD&
        '
        On Error Resume Next
        Codigo$ = TRIM$(HOJA.Cells(Fila&, 1))
        Ci% = CODINT%(Codigo$)
        COFAMILIA$ = UCase$(TRIM$(HOJA.Cells(Fila&, 2)))
        On Error GoTo 0
        If Ci% < 1 Then
            Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente.\Dar de Alta y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        U& = U& + 1
        If Codigo$ <> "" Then
           Call REPLA(X$, COFAMILIA$, 1, 2)
           Call GRAREG("!TAFAMIL", X$, U&)
        End If
        Call ACTUREGISTRO("MASTER", "CFAMILIA", "Codint = " & Ci%, COFAMILIA$, REGAF&, "NOMESS")
20 IMPORTACION = True
   Next Fila&
   '
   Libro.Close
   objexcel.Quit
   Call SETULTREG("!TAFAMIL", U&)
   Call TRACENTRAL("TAFAMIL", "")
   Call SETULTREG("TAFAMIL", U&)

   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")

99 Command8.Enabled = False
    '
End Sub






 Sub Command9_Click()
    Command9.Enabled = False
    Call LIMASTER(1)
    Command9.Enabled = True
End Sub
'
 Sub Command7_Click()
    Command7.Enabled = False
    Call LIMASTER(3)
    Command7.Enabled = True
End Sub

 Sub Command15_Click()
    Command15.Enabled = False
    Call LIMASTER(2)
    Command15.Enabled = True
End Sub
'
 Sub Command2_Click()
   Command2.Enabled = False
   Call COMUNI("Solo Disponible como\Consulta por Pantalla")
   Call Command3_Click
   Command2.Enabled = True
End Sub

 Sub Command3_Click()
   Command3.Enabled = False
   Call APERBASDAT("MASTER")
   '
   Screen.MousePointer = 11
   LITBAJ07.List1.Clear
   '
   'Set MasterBaja = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT FROM Master WHERE STATUS < 0 ORDER BY CODIGO ASC;", dbOpenSnapshot)
   SQLX$ = "SELECT CODIGO,DESCRIPCION,CODINT "
   SQLX$ = SQLX$ + " FROM Master "
   SQLX$ = SQLX$ + " WHERE STATUS <0 ORDER BY CODIGO ASC;"
   Dim MasterBaja As ADODB.Recordset
   Set MasterBaja = New ADODB.Recordset
   MasterBaja.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   On Error Resume Next
   MasterBaja.MoveLast
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "No hay Items de Baja\que Puedan Consultarse o Restaurarse.")
     GoTo 99
   End If
   On Error GoTo 0
   '
   CAREC% = MasterBaja.RecordCount
   If CAREC% < 1 Then GoTo 99
   '
   LONCO% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LONCO% < 16 Or LONCO% > 24 Then LONCO% = 16
   '
   FIN& = CAREC%
   REBLA$ = REGBLAN$("ECOMAST")
   JJ& = 0
   MasterBaja.MoveFirst
   '
10 Call TRACE(20, JJ&, FIN&)
   TTXX$ = Space$(80)
   Call REPLA(TTXX$, Left$(SAFETEXT$(MasterBaja!Codigo), LONCO%), 1, LONCO%)
   Call REPLA(TTXX$, SAFETEXT$(MasterBaja!Descripcion), LONCO% + 3, 46)
   Call REPLA(TTXX$, Str$(MasterBaja!CODINT), 73, 8)
   LITBAJ07.List1.AddItem TTXX$
   JJ& = JJ& + 1
   '
   If JJ& < CAREC% Then
     MasterBaja.MoveNext
     GoTo 10
   End If
   '
   On Error Resume Next
   MasterBaja.Close
   On Error GoTo 0
   Set MasterBaja = Nothing
   Screen.MousePointer = 1
   LITBAJ07.Show 1
   Call GENECOMAS
   '
99 Screen.MousePointer = 1
   Command3.Enabled = True
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   Call GELICOREP(2)
   Command4.Enabled = True
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   Call GELICOREP(1)
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") < 1 Then
       Call OPNOIN("Análisis Costos de Reposición")
     Else
       CONTROCOS.Show 1
   End If
   Command6.Enabled = True
   '
End Sub

Sub LIMASTER(MODO%)
    '
    HOO% = HOY(HO$)
    If MODO% = 0 Then
      MODO% = ALTERNA%("Listado de Códigos y Atributos\Código y Descripción para Inventarios\Stock Mínimo y Lote Reposición")
      If MODO% < 1 Or MODO% > 3 Then
        Exit Sub
      End If
    End If
    '
    LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
    '
    CI1% = PRICOD%
    CI2% = ULTCOD%
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
10  RACOD$ = OBJPLA$("DESHASCOD3", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = TRIM$(ECOARCH$("TATIMAT", Chr$(TMT%)))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    GRC% = CVI(Mid$(RACOD$, 6, 2))
    GRUCO$ = TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRC%)))
    If GRUCO$ = "" Then
        GRC% = 0
        Call REPLA(RACOD$, MKI$(0), 6, 2)
        GRUCO$ = "Sin Clasificación"
    End If
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    If MODO% = 1 Then
        FILEKILL LUDAT$ + "LIMASEL.DAT"
        Call HEADERS("LIMASEL", "")
        Call HEADERS("LIMASEL", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("LIMASEL", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMASEL", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
      ElseIf MODO% = 2 Then
        FILEKILL LUDAT$ + "LIMAINV.DAT"
        Call HEADERS("LIMAINV", "")
        Call HEADERS("LIMAINV", "Tipo / Grupo de Material: " + TIMAT$ + " / " + GRUCO$)
        Call HEADERS("LIMAINV", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMAINV", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
        Call HEADERS("LIMAINV", "Stock Nominal al: " + HO$ + " - " + Time$ + " hs.")
      ElseIf MODO% = 3 Then
        FILEKILL LUDAT$ + "LIMALOT.DAT"
        Call HEADERS("LIMALOT", "")
        Call HEADERS("LIMALOT", "Material: " + TRIM$(TIMAT$) + " (" + TRIM$(COD1$) + " - " + TRIM$(COD2$) + ")")
      Else
        Screen.MousePointer = 1
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    'Consulta sql de filtro
    SQLX$ = "SELECT * FROM Master "
    SQLX$ = SQLX$ + "WHERE CODIGO>='" & COD1$ & "' "
    SQLX$ = SQLX$ + "AND CODIGO<='" & COD2$ & "' "
    SQLX$ = SQLX$ + "AND STATUS>=1 "
    SQLX$ = SQLX$ + "AND (" & TMT% & " =0 OR TIMATE = " & TMT% & ") "
    SQLX$ = SQLX$ + "AND (" & GRC% & " =0 OR GRUCOVEN = " & GRC% & ") "
    SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
    ' Seteo el recordset
    Screen.MousePointer = 11
    'Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
    Dim MasterTemp As ADODB.Recordset
    Set MasterTemp = New ADODB.Recordset
    MasterTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    J& = 0: FIN& = NUINV%
    With MasterTemp
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 99
       End If
       
       registros& = .RecordCount
       FIN& = registros&
       KU& = 0
       '
       .MoveFirst
       Do While (.EOF = False)
          KU& = KU& + 1
          Call TRACE(20, KU&, FIN&)
          Ci% = !CODINT
          If MODO% = 1 Then
                'Genera el listado de Codigos y Atributos
                ML$ = REGBLAN$("LIMASEL")
                Call REPLA(ML$, MKI$(Ci%), 1, 2)
                Call REPLA(ML$, Chr$(TIMATE%(Ci%)), 3, 1)
                Call REPLA(ML$, Chr$(CRITREP%(Ci%)), 4, 1)
                Call REPLA(ML$, Chr$(MONECOSTO%(Ci%)), 5, 1)
                Call REPLA(ML$, MKI$(STOMIN(Ci%)), 15, 2)
                Call REPLA(ML$, MKI$(LOTESTAN(Ci%)), 17, 2)
                Call REPLA(ML$, Unimed$(Ci%), 19, 2)
                Call REPLA(ML$, MKS$(COSUNI(Ci%)), 21, 4)
                Call REPLA(ML$, MKI$(FECOSTO%(Ci%)), 25, 2)
                Call REPLA(ML$, MKI$(PROVHABI(Ci%)), 33, 2)
                Call REPLA(ML$, MKI$(GRUCOFAB(Ci%)), 35, 2)
                JJ& = JJ& + 1: Call GRAREG("LIMASEL", ML$, JJ&)
             ElseIf MODO% = 2 Then
                'Genera Listado de inventario
                ML$ = MKI$(0) + "  " + MKS$(0) + Space$(56)
                'Call REGAPP("LIMAINV", ML$)
                ML$ = REGBLAN$("LIMAINV")
                Call REPLA(ML$, MKI$(Ci%), 1, 2)
                Call REPLA(ML$, Unimed$(Ci%), 3, 2)
                Call REPLA(ML$, MKS$(STOCKACT(Ci%)), 5, 4)
                JJ& = JJ& + 1: Call GRAREG("LIMAINV", ML$, JJ&)
             ElseIf MODO% = 3 Then
                'Genera Listado de maestro por lote
                ML$ = REGBLAN$("LIMALOT")
                Call REPLA(ML$, MKI$(Ci%), 1, 2)
                Call REPLA(ML$, MKS$(STOMIN(Ci%)), 3, 4)
                Call REPLA(ML$, MKS$(LOTESTAN(Ci%)), 7, 4)
                Call REPLA(ML$, MKS$(LOREPOMA(Ci%)), 11, 4)
                Call REPLA(ML$, MKS$(COSUNI(Ci%)), 15, 4)
                Call REPLA(ML$, MKI$(FECOSTO%(Ci%)), 19, 2)
                JJ& = JJ& + 1: Call GRAREG("LIMALOT", ML$, JJ&)
          End If
          '
       .MoveNext
145    Loop
    End With
    Set MasterTemp = Nothing
    '
149 FIECO% = 0
    If MODO% = 1 Then
        Call SETULTREG("LIMASEL", JJ&)
        Call CIERRARCH("LIMASEL")
        Call FINAL("*LIMASEL")
      ElseIf MODO% = 2 Then
        Call SETULTREG("LIMAINV", JJ&)
        Call CIERRARCH("LIMAINV")
        Call FINAL("*LIMAINV")
      ElseIf MODO% = 3 Then
        Call SETULTREG("LIMALOT", JJ&)
        Call CIERRARCH("LIMALOT")
        Call FINAL("*LIMALOT")
    End If
    '
99  Screen.MousePointer = 1
End Sub
'
Sub GELICOREP(MODO%)
   '
   ' MODO% = 1 --> Ordenado por Proveedor
   ' MODO% = 2 --> Ordenado por Articulo
   '
   Screen.MousePointer = 11
   'Consulta sql de filtro
   SQLX$ = "SELECT * FROM Master "
   SQLX$ = SQLX$ + "WHERE STATUS>=1 "
   If MODO% = 1 Then
      SQLX$ = SQLX$ + "AND PROVHABI > 0 ORDER BY PROVHABI ASC;"
    Else
      SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
   End If
   '
   ' Seteo el recordset
   Call APERBASDAT("MASTER")
   'Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim MasterTemp As ADODB.Recordset
   Set MasterTemp = New ADODB.Recordset
   MasterTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With MasterTemp
       ingreso$ = Now
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 99
       End If
       FIN& = .RecordCount
       '
       On Error GoTo 0
       .MoveFirst
       JJ& = 0
       REBLA$ = REGBLAN$("LICOREP")
       Call BLOQARCH("LICOREP")
       Do While (.EOF = False)
          Call TRACE(20, JJ&, FIN&)
          On Error Resume Next
          CIXI% = !CODINT
          PHHI% = !PROVHABI
          UMEDIDAX$ = SAFETEXT$(!UMEDIDA)
          MONECOSTOX% = XVALO(!Monecos)
          COSUNIX = XVALO(!COSUNI)
          FECOST$ = Format(SAFETEXT$(!FECOSTO), "dd/mm/yy")
          FECOSTOX% = FECHANUM(FECOST$)
          
          On Error GoTo 0
          '
          ZZ$ = REBLA$
          Call REPLA(ZZ$, MKI$(PHHI%), 1, 2)
          Call REPLA(ZZ$, MKI$(CIXI%), 3, 2)
          Call REPLA(ZZ$, UMEDIDAX$, 5, 2) 'UMEDIDA
          Call REPLA(ZZ$, MKI$(MONECOSTOX%), 7, 2) 'Monecos
          Call REPLA(ZZ$, MKS$(COSUNIX), 9, 4) 'COSUNI
          Call REPLA(ZZ$, MKI$(FECOSTOX%), 13, 2) 'FECOSTO 'CVINT(CVDAT(
          Call REPLA(ZZ$, MKI$(PHHI%), 15, 2)
          JJ& = JJ& + 1
          Call GRAREG("LICOREP", ZZ$, JJ&)
          'CODIGO ANTERIOR SE CAMBIA PARA QUITAR LAS FUNCIONES
'          ZZ$ = REBLA$
'          Call REPLA(ZZ$, MKI$(PHHI%), 1, 2)
'          Call REPLA(ZZ$, MKI$(CIXI%), 3, 2)
'          Call REPLA(ZZ$, UNIMED$(CIXI%), 5, 2) 'UMEDIDA
'          Call REPLA(ZZ$, MKI$(MONECOSTO%(CIXI%)), 7, 2) 'Monecos
'          Call REPLA(ZZ$, MKS$(COSUNI(CIXI%)), 9, 4) 'COSUNI
'          Call REPLA(ZZ$, MKI$(FECOSTO%(CIXI%)), 13, 2) 'FECOSTO 'CVINT(CVDAT(
'          Call REPLA(ZZ$, MKI$(PHHI%), 15, 2)
'          JJ& = JJ& + 1
'          Call GRAREG("LICOREP", ZZ$, JJ&)
          '
       .MoveNext
       Loop
       '
   End With
   EGRESO$ = Now
   Set MasterTemp = Nothing
   Call SETULTREG("LICOREP", JJ&)
   Call CIERRARCH("LICOREP")
   '
   Call CONECRUN("*LICOREP", "Listado Costos de Reposición por Proveedor")
   '
99 Screen.MousePointer = 1
   End Sub


Private Sub Form_Load()
   If InStr(EMPREAC$, "ACHERN") > 0 Then
     Command12.Visible = True
   End If
   If InStr(EMPREAC$, "GABAL") > 0 Then
     'Command13.Visible = True
     Command14.Visible = True
     'FECHA DE ULTIMO CONTROL DE PERIDO DE STOCKS MINIMO
     Command16.Visible = True
   End If
   '
   'Call APLICACIONSKINS(Me, Picture6)
End Sub

'
'

Function UltimoRegistro(COLUMNA) As Double
   Dim valor As Variant
   Dim CantFilasVacias&
   CantFilasVacias& = 0
   UltFila& = 0
   For Fila& = 2 To HOJA.Rows.Count
      If CantFilasVacias >= 100 Then
         GoTo 10
      End If
         valor = TRIM$(HOJA.Cells(Fila&, COLUMNA))
         If valor = "" Then 'Si no hay valor
            If UltFila& = 0 Then 'Si no hay registro de ultima fila se registra, si no se mantiene la almacenada
               UltFila& = Fila& - 1
            End If
            CantFilasVacias& = CantFilasVacias& + 1
           Else
            CantFilasVacias& = 0
            UltFila& = 0
          End If
   Next Fila&
10 UltimoRegistro = UltFila&
End Function


Private Sub Label7_Click()

End Sub
