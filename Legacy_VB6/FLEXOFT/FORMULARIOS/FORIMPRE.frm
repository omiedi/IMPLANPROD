VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FORIMPRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Maestro"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
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
      Picture         =   "FORIMPRE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Importación"
      Top             =   945
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\LISTA 104 a Migrar.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Sheet1$"
      Top             =   105
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
      Picture         =   "FORIMPRE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORIMPRE.frx":0454
      Height          =   7470
      Left            =   0
      OleObjectBlob   =   "FORIMPRE.frx":0464
      TabIndex        =   1
      Top             =   315
      Width           =   10830
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aro Plastico"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   7
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resorte"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   6
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carcaza"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   5
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acelerante"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   4
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mezcla"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   3
      Top             =   315
      Width           =   1380
   End
End
Attribute VB_Name = "FORIMPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Close
   Hide
End Sub
Private Sub Command4_Click()
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 50) * DBGrid1.RowHeight
   '
   JJ& = 0
   REGX& = 0: REGY& = 0: NCIMAX% = 0
   '
' IMPORTACION PRECIOS
   '
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   '
   JJL& = 0
   '
   For U& = 7 To URECORD&
       DBGrid1.Row = U&
       DBGrid1.Col = 0
       CODIAR$ = TRIM$(DBGrid1.TEXT)
       CII% = CODINT%(CODIAR$)
       If CII% > 0 Then
         DBGrid1.Col = 1
         DESAR$ = TRIM$(DBGrid1.TEXT)
         NORIG$ = ""
         If Val(Left$(DESAR$, 6)) > 1000 Then
           If Mid$(DESAR$, 7, 1) = " " Then
             NORIG$ = TRIM$(Left$(DESAR$, 6))
             'DESAR$ = TRIM$(Mid$(DESAR$, 8))
           End If
         End If
         DBGrid1.Col = 2
         PREUN# = Val(DBGrid1.TEXT)
25       Z$ = REGBLAN$("LIPRE001")
         Call REPLA(Z$, MKI$(CII%), 1, 2)
         Call REPLA(Z$, DESAR$, 3, 42)
         Call REPLA(Z$, NORIG$, 46, 16)
         Call REPLA(Z$, MKD$(PREUN#), 62, 8)
         JJL& = JJL& + 1
         Call GRAREG("LIPRE001", Z$, JJL&)
       End If
   Next U&
   '
   Call SETULTREG("LIPRE001", JJL&)
   Call CIERRARCH("LIPRE001")
   MsgBox "FIN"
   Call CIERRARCH("*.*")
   '
End Sub

