VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CONSULTA1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GASTOS TOTALES POR CENTRO DE COSTOS CORRESPONDIENTE AL AÑO: "
   ClientHeight    =   6675
   ClientLeft      =   1050
   ClientTop       =   870
   ClientWidth     =   15855
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   7455
      Begin VB.ComboBox CMBFILTRO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   160
         Width           =   3135
      End
      Begin VB.TextBox TXTBUSCAR 
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
         Left            =   720
         TabIndex        =   12
         Top             =   160
         Width           =   2535
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3360
         TabIndex        =   15
         Top             =   160
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   160
         Width           =   585
      End
   End
   Begin VB.TextBox TXTDATO3 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   1440
      Width           =   6000
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Dato a Buscar 3"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox TXTDATO2 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   6000
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Dato a Buscar 2"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox TXTDATO1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   6000
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Dato a Buscar 1"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar Consulta"
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
      Left            =   5280
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3675
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   3675
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "CONSULTA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
     Command1.Enabled = False
     GRID.Rows = 1
     Dim VARS()
     If TRIM$(TXTDATO1.TEXT) <> "" Then BUSCAR1% = 1: DATOX$ = TXTDATO1.TEXT
     If TRIM$(TXTDATO2.TEXT) <> "" Then BUSCAR2% = 1
     If TRIM$(TXTDATO3.TEXT) <> "" Then BUSCAR3% = 1
     If BUSCAR1% < 1 Then
        Call COMUNI("Falta Ingresar Datos")
        Exit Sub
     End If
     
If EXISTE_SP%("SearchAllTables") > 0 Then
  SQLX$ = "Alter "
Else
  SQLX$ = "Create "
  EJECUTA_DIRECTO% = 1
End If
SQLX$ = SQLX$ + " PROCEDURE SearchAllTables "
SQLX$ = SQLX$ + "( "
SQLX$ = SQLX$ + "@SearchStr nvarchar(100) "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + "AS "
SQLX$ = SQLX$ + "BEGIN "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "CREATE TABLE #Results (ColumnName nvarchar(370), ColumnValue nvarchar(3630)) "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "SET NOCOUNT ON "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "DECLARE @TableName nvarchar(256), @ColumnName nvarchar(128), @SearchStr2 nvarchar(110) "
SQLX$ = SQLX$ + "SET  @TableName = '' "
SQLX$ = SQLX$ + "SET @SearchStr2 = QUOTENAME('%' + @SearchStr + '%','''') "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "WHILE @TableName IS NOT NULL "
SQLX$ = SQLX$ + "BEGIN "
SQLX$ = SQLX$ + "SET @ColumnName = '' "
SQLX$ = SQLX$ + "SET @TableName = "
SQLX$ = SQLX$ + "( "
SQLX$ = SQLX$ + "SELECT MIN(QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME)) "
SQLX$ = SQLX$ + "FROM    INFORMATION_SCHEMA.TABLES "
SQLX$ = SQLX$ + "WHERE       TABLE_TYPE = 'BASE TABLE' "
SQLX$ = SQLX$ + "AND QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME) > @TableName "
SQLX$ = SQLX$ + "AND OBJECTPROPERTY( "
SQLX$ = SQLX$ + "OBJECT_ID( "
SQLX$ = SQLX$ + "QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME) "
SQLX$ = SQLX$ + "), 'IsMSShipped' "
SQLX$ = SQLX$ + ") = 0 "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "WHILE (@TableName IS NOT NULL) AND (@ColumnName IS NOT NULL) "
SQLX$ = SQLX$ + "BEGIN "
SQLX$ = SQLX$ + "SET @ColumnName = "
SQLX$ = SQLX$ + "( "
SQLX$ = SQLX$ + "SELECT MIN(QUOTENAME(COLUMN_NAME)) "
SQLX$ = SQLX$ + "FROM    INFORMATION_SCHEMA.COLUMNS "
SQLX$ = SQLX$ + "WHERE       TABLE_SCHEMA    = PARSENAME(@TableName, 2) "
SQLX$ = SQLX$ + "AND TABLE_NAME  = PARSENAME(@TableName, 1) "
SQLX$ = SQLX$ + "AND DATA_TYPE IN ('char', 'varchar', 'nchar', 'nvarchar') "
SQLX$ = SQLX$ + "AND QUOTENAME(COLUMN_NAME) > @ColumnName "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "IF @ColumnName IS NOT NULL "
SQLX$ = SQLX$ + "BEGIN "
SQLX$ = SQLX$ + "INSERT INTO #Results "
SQLX$ = SQLX$ + "EXEC "
SQLX$ = SQLX$ + "( "
SQLX$ = SQLX$ + "'SELECT ''' + @TableName + '.' + @ColumnName + ''', LEFT(' + @ColumnName + ', 3630) "
SQLX$ = SQLX$ + "FROM ' + @TableName + ' (NOLOCK) ' + "
SQLX$ = SQLX$ + "' WHERE ' + @ColumnName + ' LIKE ' + @SearchStr2 "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + "END "
SQLX$ = SQLX$ + "END "
SQLX$ = SQLX$ + "END "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "SELECT ColumnName, ColumnValue FROM #Results "
SQLX$ = SQLX$ + "END "
If EJECUTA_DIRECTO% > 0 Then
  FLEXCONN.Execute (SQLX$)
Else
  If CAMBIOSP("SearchAllTables", SQLX$) Then
   FLEXCONN.Execute (SQLX$)
  End If
End If


10 SQLX$ = " EXECUTE SearchAllTables  '" & TRIM$(UCase$(DATOX$)) & "'"
'   SQLX$ = SQLX$ + "," & CodiEmp%
   Set RS = FLEXCONN.Execute(SQLX$)
   U& = 0
   With RS
      Do While Not .EOF
        For I& = 0 To RS.Fields.Count - 1
             VALOR$ = RS.Fields(I&).Value
             CACAM& = RS.Fields.Count
             ReDim VARS(CACAM&)
             For U& = 1 To CACAM&
               VARS(U&) = RS.Fields(U& - 1).Value
             Next U&
             GoSub AGREGARFILA
        Next I&
        .MoveNext
      Loop
      GoSub AGREGARFILA
   End With
   Command1.Enabled = True
   If BUSCAR2% = 1 Then
      BUSCAR2% = 0
      DATOX$ = TXTDATO2.TEXT
      GoTo 10
   ElseIf BUSCAR3% = 1 Then
      BUSCAR3% = 0
      DATOX$ = TXTDATO3.TEXT
      GoTo 10
   End If
   Exit Sub
   
AGREGARFILA:
        GRID.Rows = GRID.Rows + 1
        For j& = 1 To CACAM&
            TEXTOX$ = TRIM$(UCase$(VARS(j&)))
            If Left$(TEXTOX$, 5) = "[DBO]" Then TEXTOX$ = Mid$(TEXTOX$, 6)
            GRID.TextMatrix(GRID.Rows - 1, j&) = TEXTOX$
        Next j&
Return
End Sub
Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub

Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For I& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, I&) = ""
    Next I&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For I& = 1 To GRID.Rows - 1
      For j& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(I&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(I&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next I&
    '
    MSF_2.Redraw = True
    
    '
End Sub
Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub
Private Sub Form_Load()
     CAMPOS$ = "Tabla-Campo/A64;Dato/A64"
     Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me, 1)
     Call CARGA_ENCABEZADO(Me.MSF_2, CAMPOS$, Me, 1)
'    Call ABRECONEXION
'    Combo1.Clear
'    Dim AUX As ADODB.Recordset
'    Set AUX = New ADODB.Recordset
'22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
'    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    Do While Not AUX.EOF
'       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
'       Combo1.AddItem UCase$(AUX!Name)
'29     AUX.MoveNext
'    Loop

End Sub

Private Sub TXTDATO1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TXTDATO2.SetFocus
   End If
End Sub


Private Sub TXTDATO2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TXTDATO3.SetFocus
   End If

End Sub


Private Sub TXTDATO3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Command1.SetFocus
   End If

End Sub


