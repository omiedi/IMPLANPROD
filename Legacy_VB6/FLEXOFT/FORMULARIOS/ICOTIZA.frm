VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ICOTIZA 
   Caption         =   "Importar cotizacion dolar"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   5145
      Picture         =   "ICOTIZA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   5145
      Picture         =   "ICOTIZA.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Iniciar Lectura de Archivo"
      Top             =   3150
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "C:\Mis documentos\COTIZACION DOLAR.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Hoja1$"
      Top             =   3255
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ICOTIZA.frx":0454
      Height          =   3840
      Left            =   105
      OleObjectBlob   =   "ICOTIZA.frx":0464
      TabIndex        =   0
      Top             =   210
      Width           =   4890
   End
End
Attribute VB_Name = "ICOTIZA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Exit Sub  '(IMPORTADOR DE COTIZACION DOLAR)
   Data1.Refresh
   Data1.Recordset.MoveFirst
   '
   URECORD& = Data1.Recordset.RecordCount
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   AGREGA% = 0
   '
   For U& = 1 To URECORD&
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     FECOT$ = Format$(DBGrid1.TEXT, "dd/mm/yy")
     '
     DBGrid1.Col = 1
     COTIZ# = 0
     On Error Resume Next
     COTIZ# = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     '
     FFX% = FECHANUM(FECOT$)
     If FFX% > 0 Then
       If COTIZ# >= 1 Then
         ZZ$ = REGBLAN$("CODIADOL")
         Call REPLA(ZZ$, MKI$(FECHANUM(FECOT$)), 1, 2)
         Call REPLA(ZZ$, MKD$(COTIZ#), 3, 8)
         '
         For JJ& = 1 To ULTREG&("CODIADOL")
           XX$ = REGLEIDO$("CODIADOL", JJ&)
           If CVI(Left$(XX$, 2)) = FFX% Then
             Call GRAREG("CODIADOL", ZZ$, JJ&)
             GoTo 89
           End If
         Next JJ&
         Call REGAPP("CODIADOL", ZZ$)
         AGREGA% = 1
       End If
     End If
89 Next U&
   '
   Call CIERRARCH("CODIADOL")
   If AGREGA% > 0 Then
     Call FILESORT("CODIADOL", "", 1, 1, "ASC")
     Call CIERRARCH("CODIADOL")
     Call BAJALDISCO
   End If
   '
   Call COMUNI("Fin Importación")
   '
End Sub

Private Sub Command2_Click()
   Hide
End Sub
